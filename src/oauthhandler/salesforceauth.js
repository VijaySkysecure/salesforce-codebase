const axios = require("axios");
const querystring = require("querystring");
const config = require("./config");
const { container: containerPromise } = require("./cosmos");

function getAuthUrl(teamsChatId) {
  const params = {
    client_id: config.outlookClientId,
    response_type: "code",
    redirect_uri: config.outlookRedirectUri,
    response_mode: "query",
    scope: config.outlookScopes,
    state: teamsChatId,
  };

  return `${config.outlookAuthorityUrl}/${config.outlookTenantId}/oauth2/v2.0/authorize?${querystring.stringify(params)}`;
}

async function exchangeCodeForToken(authCode) {
  const tokenEndpoint = `${config.outlookAuthorityUrl}/${config.outlookTenantId}/oauth2/v2.0/token`;

  const body = querystring.stringify({
    client_id: config.outlookClientId,
    scope: config.outlookScopes,
    code: authCode,
    redirect_uri: config.outlookRedirectUri,
    grant_type: "authorization_code",
    client_secret: config.outlookClientSecret,
  });

  const response = await axios.post(tokenEndpoint, body, {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  return response.data;
}

async function storeOutlookToken(teamsChatId, tokenResponse) {
  const container = await containerPromise;

  const record = {
    id: `outlook-${teamsChatId}`,
    teamsChatId,
    accessToken: tokenResponse.access_token,
    refreshToken: tokenResponse.refresh_token,
    expiresAt: Date.now() + tokenResponse.expires_in * 1000,
    partitionKey: `outlook-${teamsChatId}`,
  };

  await container.items.upsert(record);
  return record;
}

async function outLookRefreshAccessToken(teamsChatId, refreshToken) {
  try {
    const { outlookClientId, outlookClientSecret } = config
    const container = await containerPromise;
    const response = await axios.post(
      "https://login.microsoftonline.com/common/oauth2/v2.0/token",
      null,
      {
        params: {
          grant_type: "refresh_token",
          client_id: outlookClientId,
          client_secret: outlookClientSecret,
          refresh_token: refreshToken,
        },
      }
    );

    const { access_token, expires_in } = response.data;

    if (!access_token || !expires_in) {
      throw new Error("Invalid refresh token response");
    }

    const tokenData = {
      id: teamsChatId,
      teamsChatId,
      userId: (await container.item(teamsChatId, teamsChatId).read()).resource?.userId,
      accessToken: access_token,
      refreshToken,
      expiresAt: Date.now() + expires_in * 1000,
      partitionKey: teamsChatId,
    };

    await container.items.upsert(tokenData);
    console.log(`Refreshed token for teamsChatId: ${teamsChatId}`);
    return access_token;
  } catch (error) {
    console.error(`Error refreshing token for teamsChatId ${teamsChatId}:`, error.message);
    throw error;
  }
}

async function getOutlookToken(teamsChatId) {
  try {

    const container = await containerPromise;
    const { resource: token } = await container.item(`outlook-${teamsChatId}`, `outlook-${teamsChatId}`).read();

    if (!token) {
      return false
    }
    if (Date.now() > token.expiresAt) {
      const accessToken = await outLookRefreshAccessToken(teamsChatId, token.refreshToken)
      return accessToken;
    }

    return token.accessToken;
  } catch (error) {
    console.error("Error retrieving Outlook token:", error.message)
    return false;
  }
}

module.exports = {
  getAuthUrl,
  exchangeCodeForToken,
  storeOutlookToken,
  getOutlookToken,
};