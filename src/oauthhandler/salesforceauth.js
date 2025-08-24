const axios = require("axios");
const querystring = require("querystring");
const config = require("../config");
const { container: containerPromise } = require("../cosmos");
const { MessageFactory } = require("botbuilder");
const { htmlToText } = require("html-to-text");


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


async function initializeConversationStateOutlook(context, state, teamsChatId) {
  try {
    // Initialize conversation state if not present
    if (!state.conversation) {
      state.conversation = {};
    }
    
    const token = await getOutlookToken(teamsChatId);
    if (token && token !== false) {
      state.conversation.isOutlookAuthenticated = true;
      state.conversation.userId = context.activity.from.id;
      console.log(`Outlook authentication successful for teamsChatId: ${teamsChatId}`);
      return true;
    } else {
      state.conversation.isOutlookAuthenticated = false;
      console.log(`Outlook authentication failed for teamsChatId: ${teamsChatId}`);
      return false;
    }
  } catch (error) {
    console.error("Error initializing Outlook conversation state:", error);
    state.conversation.isOutlookAuthenticated = false;
    return false;
  }
}


async function getRecentEmails(context, state, limit) {
  try {
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const token = await getOutlookToken(teamsChatId);
    if (!token) {
      state.conversation.isOutlookAuthenticated = false;
      await context.sendActivity(
        MessageFactory.text(
          `[GetRecentEmails] 🔒 You need to authenticate with Outlook first. Please use the \`/outlook\` command to login.`
        )
      );
      return { status: "error", message: "User authentication required" };
    }
    state.conversation.isOutlookAuthenticated = true;
    state.conversation.userId = userId;
    console.log("Fetching recent emails from Outlook...");
    let retries = 3;
    while (retries > 0) {
      try {
        const response = await axios.get(
          `https://graph.microsoft.com/v1.0/me/messages?$orderby=receivedDateTime desc&$top=${limit}`,
          {
            headers: { Authorization: `Bearer ${token}` },
          }
        );
        console.log("Raw API response for recent emails:", JSON.stringify(response.data, null, 2));
        return {
          status: "success",
          data: response.data.value || [],
          message: response.data.value?.length > 0 ? "Emails found" : "No emails found"
        };
      } catch (error) {
        if (error.response?.status === 429 && retries > 0) {
          const delay = Math.pow(2, 3 - retries) * 1000;
          console.log(`Rate limit hit, retrying after ${delay}ms...`);
          await new Promise((resolve) => setTimeout(resolve, delay));
          retries--;
          continue;
        }
        throw error;
      }
    }
    throw new Error("Max retries reached for rate limit");
  } catch (error) {
    console.error(`Error fetching recent emails:`, {
      message: error.message,
      status: error.response?.status,
      data: error.response?.data
    });
    if (error.response?.status === 401) {
      state.conversation.isOutlookAuthenticated = false;
      await context.sendActivity(
        MessageFactory.text(`[GetRecentEmails] 🔑 Your Outlook session has expired. Please use \`/outlook\` to login again.`)
      );
      return { status: "error", message: "Token expired" };
    }
    const errorMessage = error.response?.data?.error?.message || error.message;
    return { status: "error", message: errorMessage };
  }
}


module.exports = {
  getAuthUrl,
  exchangeCodeForToken,
  storeOutlookToken,
  getOutlookToken,
  initializeConversationStateOutlook,
  getRecentEmails,
};