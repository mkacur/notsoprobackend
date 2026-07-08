  import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";
import cors from "cors";
import { v4 as uuidv4 } from "uuid";

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 10000;

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

// Your SharePoint site path
const sitePath = process.env.SHAREPOINT_SITE_PATH;

// ============================================================
// 1. Get Access Token (Client Credentials Flow)
// ============================================================
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append("client_id", clientId);
  params.append("client_secret", clientSecret);
  params.append("scope", "https://graph.microsoft.com/.default");
  params.append("grant_type", "client_credentials");

  const response = await fetch(url, {
    method: "POST",
    body: params,
  });

  const data = await response.json();
  return data.access_token;
}

// ============================================================
// 2. Fetch ALL pages of a SharePoint list
// ============================================================
async function getListItems(listName) {
  const token = await getAccessToken();

  // Step A: Get siteId
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${sitePath}`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  const siteData = await siteRes.json();
  const siteId = siteData.id;

  // Step B: Get listId
  const listRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  const listData = await listRes.json();
  const listId = listData.id;

  // Step C: Fetch all pages
  let items = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$top=999`;

  while (nextUrl) {
    const pageRes = await fetch(nextUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });

    const pageData = await pageRes.json();
    if (pageData.value) items = items.concat(pageData.value);

    nextUrl = pageData["@odata.nextLink"] || null;
  }

  return { value: items };
}

// ============================================================
// 3. Update a SharePoint list item
// ============================================================
async function updateListItem(listName, itemId, fields) {
  const token = await getAccessToken();

  // Step A: Get siteId
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${sitePath}`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  const siteData = await siteRes.json();
  const siteId = siteData.id;

  // Step B: Get listId
  const listRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );
  const listData = await listRes.json();
  const listId = listData.id;

  // Step C: PATCH
  const updateRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(fields),
    }
  );

  return updateRes.json();
}

// ============================================================
// 4. Basic GET Routes (unchanged)
// ============================================================
app.get("/api/divisions", async (req, res) => {
  try {
    const data = await getListItems("Divisions");
    res.json(data);
  } catch (err) {
    console.error("Error fetching Divisions:", err);
    res.status(500).json({ error: "Failed to fetch Divisions" });
  }
});

app.get("/api/teams", async (req, res) => {
  try {
    const data = await getListItems("Teams");
    res.json(data);
  } catch (err) {
    console.error("Error fetching Teams:", err);
    res.status(500).json({ error: "Failed to fetch Teams" });
  }
});

app.get("/api/games", async (req, res) => {
  try {
    const data = await getListItems("Games");
    res.json(data);
  } catch (err) {
    console.error("Error fetching Games:", err);
    res.status(500).json({ error: "Failed to fetch Games" });
  }
});

app.get("/api/admin", async (req, res) => {
  try {
    const data = await getListItems("Admin");
    res.json(data);
  } catch (err) {
    console.error("Error fetching Admin:", err);
    res.status(500).json({ error: "Failed to fetch Admin" });
  }
});

// ============================================================
// 5. PATCH Routes (unchanged)
// ============================================================
app.patch("/api/games/:id", async (req, res) => {
  try {
    const result = await updateListItem("Games", req.params.id, req.body);
    res.json({ success: true, result });
  } catch (err) {
    console.error("Error updating game:", err);
    res.status(500).json({ error: "Failed to update game" });
  }
});

app.patch("/api/teams/:id", async (req, res) => {
  try {
    const result = await updateListItem("Teams", req.params.id, req.body);
    res.json({ success: true, result });
  } catch (err) {
    console.error("Error updating team:", err);
    res.status(500).json({ error: "Failed to update team" });
  }
});

// ============================================================
// 6. NEW RESET SYSTEM (background job + polling)
// ============================================================

// In-memory job store
const resetJobs = {};

// Fetch admin password
async function getAdminPassword() {
  const admin = await getListItems("Admin");
  if (!admin.value.length) throw new Error("Admin list empty.");
  return admin.value[0].fields.Password;
}

// Reset Teams
async function resetTeams(job) {
  const teams = await getListItems("Teams");
  job.teamsTotal = teams.value.length;
  job.teamsDone = 0;

  for (const t of teams.value) {
    await updateListItem("Teams", t.id, {
      Wins: 0,
      Losses: 0,
      For: 0,
      Ag: 0,
      Diff: 0,
    });
    job.teamsDone++;
  }
}

// Reset Games
async function resetGames(job) {
  const games = await getListItems("Games");
  job.gamesTotal = games.value.length;
  job.gamesDone = 0;

  for (const g of games.value) {
    await updateListItem("Games", g.id, {
      ScoreA: null,
      ScoreB: null,
      Winner: null,
      Loser: null,
      Status: "Not Started",
    });
    job.gamesDone++;
  }
}

// Background job runner
async function runResetJob(jobId) {
  const job = resetJobs[jobId];
  if (!job) return;

  job.status = "running";
  job.startedAt = new Date();

  try {
    await resetTeams(job);
    await resetGames(job);

    job.status = "complete";
    job.finishedAt = new Date();
    job.message = "Reset complete.";
  } catch (err) {
    console.error("Reset job failed:", err);
    job.status = "error";
    job.finishedAt = new Date();
    job.message = err.message;
  }
}

// Start reset
app.post("/api/start-reset", async (req, res) => {
  try {
    const { password } = req.body;
    if (!password) return res.status(400).json({ error: "Password required." });

    const adminPassword = await getAdminPassword();
    if (password !== adminPassword)
      return res.status(401).json({ error: "Invalid password." });

    const jobId = uuidv4();
    resetJobs[jobId] = {
      status: "pending",
      message: "Job created.",
      teamsDone: 0,
      teamsTotal: 0,
      gamesDone: 0,
      gamesTotal: 0,
      startedAt: null,
      finishedAt: null,
    };

    runResetJob(jobId); // background

    res.json({ jobId, status: "started" });
  } catch (err) {
    console.error("Error starting reset:", err);
    res.status(500).json({ error: "Failed to start reset." });
  }
});

// Poll reset status
app.get("/api/reset-status", (req, res) => {
  const { jobId } = req.query;
  const job = resetJobs[jobId];

  if (!job) return res.status(404).json({ error: "Job not found." });

  res.json(job);
});

// ============================================================
// 7. Start Server
// ============================================================
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
