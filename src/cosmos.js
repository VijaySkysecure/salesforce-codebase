const { CosmosClient } = require("@azure/cosmos");
const config = require("./config");

const cosmosClient = new CosmosClient({
  endpoint: config.cosmosEndpoint,
  key: config.cosmosKey,
});

// Initialize database and container
async function initializeCosmos() {
  try {
    const { database } = await cosmosClient.databases.createIfNotExists({ id: config.cosmosDatabaseId });
    const { container } = await database.containers.createIfNotExists({
      id: config.cosmosContainerId,
      partitionKey: { paths: ["/teamsChatId"] },
    });
    console.log("Cosmos DB initialized successfully");
    return container;
  } catch (error) {
    console.error("Error initializing Cosmos DB:", error.message);
    throw error;
  }
}

async function storeUserToken(teamsChatId, type, tokenResponse) {
  const container = await containerPromise;
 
  const record = {
    id: teamsChatId,
    type,
    teamsChatId,
    accessToken: tokenResponse.access_token,
    refreshToken: tokenResponse.refresh_token,
    instance_url: tokenResponse.instance_url,
    signature: tokenResponse.signature,
    issuedAt: tokenResponse.issued_at,
    expiresAt: Date.now() + tokenResponse.expires_in * 1000,
  };
 
  await container.items.upsert(record);
  return record;
}

async function getUserToken (teamsChatId, type) {
  const container = await containerPromise;
  const { resource: item } = await container.item(teamsChatId, teamsChatId).read();
  if (item && item.type === type) {
    return {
      status: true,
      accessToken: item.accessToken,
      refreshToken: item.refreshToken,
      instanceUrl: item.instanceUrl,
    };
  }
  return {
    status: false  
  };
}
 

// Export initialized container
module.exports = {
  container: initializeCosmos(),
  storeUserToken,
  getUserToken
};