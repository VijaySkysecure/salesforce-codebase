// Import required packages
const express = require("express");
const axios = require("axios");

const MongoDbConnection = require("./config/mongoose");

// This agent's adapter
const adapter = require("./adapter");

// This agent's main dialog.
const app = require("./app/app");


const { storeUserToken } = require("./user")

// Create express application.
const expressApp = express();
expressApp.use(express.json());

new MongoDbConnection();

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nAgent started, ${expressApp.name} listening to`, server.address());
});

// Listen for incoming requests.
expressApp.post("/api/messages", async (req, res) => {
  // Route received a request to adapter for processing
  await adapter.process(req, res, async (context) => {
    // Dispatch to application for routing
    await app.run(context);
  });
});


// OAuth callback endpoint
expressApp.get("/salesforce/callback", async (req, res) => {
  const { code, state } = req.query;
  const { salesforceClientId, salesforceClientSecret, salesforceRedirectUri } = require("./config");
  if (!code || !state) return res.status(400).send("Missing code or state");

  // Decode and verify state
  let userId, teamsChatId;
  try {
    const decodedState = JSON.parse(Buffer.from(state, "base64url").toString());
    userId = decodedState.userId;
    teamsChatId = decodedState.teamsChatId || userId;
  } catch (error) {
    console.error("Error decoding state:", error.message);
    return res.status(400).send("Invalid state parameter");
  }
  try {
    // Exchange the code for the access token
    const { data } = await axios.post(
      "https://login.salesforce.com/services/oauth2/token",
      null,
      {
        params: {
          grant_type: "authorization_code",
          client_id: salesforceClientId,
          client_secret: salesforceClientSecret,
          redirect_uri: salesforceRedirectUri,
          code,
        },
      }
    );

    console.log("Access token received:", data);
    const result = await storeUserToken(teamsChatId, "salesforce", data);
    console.log("Token stored in Cosmos DB:", result);

    res.send(`<html><body><h2>✅ Connected. Close this tab.</h2></body></html>`);
  } catch (error) {
    console.error("Error exchanging code or storing token:", error.message);
    res.status(500).send("Error during OAuth token exchange or storage.");
  }
});

// Oauth callback for outlook


expressApp.get("/outlook/callback", async (req, res) => {
  const { code, state } = req.query;
  const { microsoftClientId, microsoftClientSecret, microsoftRedirectUrl } = require("./config");

  if (!code || !state) return res.status(400).send("Missing code or state");

  // Decode state
  let userId, teamsChatId;
  try {
    const decodedState = JSON.parse(Buffer.from(state, "base64url").toString());
    userId = decodedState.userId;
    teamsChatId = decodedState.teamsChatId || userId;
  } catch (error) {
    console.error("Error decoding state:", error.message);
    return res.status(400).send("Invalid state parameter");
  }

  try {
    // Build form body
    const body = new URLSearchParams({
      grant_type: "authorization_code",
      client_id: microsoftClientId,
      client_secret: microsoftClientSecret,
      redirect_uri: microsoftRedirectUrl,
      code,
    });

    // Send POST with body
    const { data } = await axios.post(
      "https://login.microsoftonline.com/common/oauth2/v2.0/token",
      body.toString(),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    console.log("Access token received:", data);

    await storeUserToken(teamsChatId, "outlook", data);

    res.send(`<html><body><h2>✅ Connected to Outlook. You can close this tab.</h2></body></html>`);
  } catch (error) {
    console.error("OAuth error:", error.response?.data || error.message);
    res.status(500).send("Error during OAuth token exchange or storage.");
  }
});

