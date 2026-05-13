import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const siteUrl = "https://fsavaluation.sharepoint.com/sites/notsopro";

async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams();
  params.append("grant_type", "client_credentials");
  params.append("client_id", clientId);
  params.append("client_secret", clientSecret);
  params.append("scope", "https://graph.microsoft.com/.default");

  const response = await fetch(url, {
    method: "POST",
    body: params
  });

  const data = await response.json();
  return data.access_token;
}

async function callSharePoint(listName) {
  const token = await getAccessToken();

  const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

  const response = await fetch(url, {
    headers: {
      "Authorization": `Bearer ${token}`,
      "Accept": "application/json;odata=nometadata"
    }
  });

  return response.json();
}

app.get("/api/divisions", async (req, res) => {
  const data = await callSharePoint("Divisions");
  res.json(data);
});

app.get("/api/teams", async (req, res) => {
  const data = await callSharePoint("Teams");
  res.json(data);
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
