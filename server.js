import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";
import cors from "cors";
app.use(cors());
dotenv.config();

const app = express();
const PORT = process.env.PORT || 10000;

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;

const sitePath = "fsavaluation.sharepoint.com:/sites/notsopro";

// ------------------------------------------------------------
// 1. Get Access Token (Client Credentials Flow)
// ------------------------------------------------------------
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

// ------------------------------------------------------------
// 2. Auto‑discover siteId and listId, then fetch raw items
// ------------------------------------------------------------
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

  // Step C: Get raw list items
  const itemsRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  return itemsRes.json();
}

// ------------------------------------------------------------
// 3. API Routes
// ------------------------------------------------------------
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

// ------------------------------------------------------------
// 4. Start Server
// ------------------------------------------------------------
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
