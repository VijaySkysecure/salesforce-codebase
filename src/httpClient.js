const axios = require("axios");
const { getUserToken, storeUserToken } = require("./cosmos");
const { salesforceClientId, salesforceClientSecret, salesforceRedirectUri } = require("./config");

async function refreshAccessToken(teamsChatId, refreshToken) {
    try {
        const response = await axios.post(
            "https://login.salesforce.com/services/oauth2/token",
            null,
            {
                params: {
                    grant_type: "refresh_token",
                    client_id: salesforceClientId,
                    client_secret: salesforceClientSecret,
                    refresh_token: refreshToken,
                    redirect_uri: salesforceRedirectUri
                }
            }
        );

        const newTokens = response.data;
        console.log("🔄 Refreshed Salesforce access token");

        // Store new token back in Cosmos
        await storeUserToken(teamsChatId, "salesforce", newTokens);

        return newTokens;
    } catch (error) {
        console.error("❌ Failed to refresh token:", error.response?.data || error.message);
        throw error;
    }
}

function createSalesforceClient(teamsChatId) {
    const client = axios.create();

    // Attach interceptor
    client.interceptors.response.use(
        (response) => response,
        async (error) => {
            if (error.response?.status === 401) {
                console.warn("⚠️ Salesforce token expired. Attempting refresh...");
                const { refreshToken, instanceUrl } = await getUserToken(teamsChatId, "salesforce");
                if (!refreshToken) throw new Error("No refresh token available");

                const newTokens = await refreshAccessToken(teamsChatId, refreshToken);

                // Retry original request with new access token
                error.config.headers["Authorization"] = `Bearer ${newTokens.access_token}`;
                error.config.baseURL = newTokens.instanceUrl;

                return client.request(error.config);
            }
            return Promise.reject(error);
        }
    );

    return client;
}

module.exports = { createSalesforceClient };
