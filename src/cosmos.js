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
  const container = await initializeCosmos();
  const record = {
    id: teamsChatId,
    type,
    teamsChatId,
    accessToken: tokenResponse.access_token,
    instanceUrl: tokenResponse.instance_url,
    signature: tokenResponse.signature,
    issuedAt: tokenResponse.issued_at,
    expiresAt: Date.now() + tokenResponse.expires_in * 1000,
  };
  if (tokenResponse.refresh_token) {
    record.refreshToken = tokenResponse.refresh_token;
  }
  await container.items.upsert(record);
  return record;
}

async function getUserToken (teamsChatId, type) {
  const container = await initializeCosmos();
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



// Delete item from cosmos
async function deleteUserToken(teamsChatId, type) {
  const container = await initializeCosmos();

  try {
    // Try to read the item
    const { resource: item } = await container.item(teamsChatId, teamsChatId).read();

    if (item && item.type === type) {
      await container.item(item.id, item.teamsChatId).delete();
      return { status: true, message: `✅ Deleted token for type=${type}` };
    } else {
      return { status: false, message: `⚠️ No token found for teamsChatId=${teamsChatId}, type=${type}` };
    }
  } catch (error) {
    if (error.code === 404) {
      // Item not found
      return { status: false, message: `⚠️ Token not found: teamsChatId=${teamsChatId}, type=${type}` };
    }
    console.error("❌ Error deleting token:", error);
    return { status: false, message: `❌ Error deleting token: ${error.message}` };
  }
}



 

// Export initialized container
module.exports = {
  container: initializeCosmos(),
  storeUserToken,
  getUserToken,
  deleteUserToken
};