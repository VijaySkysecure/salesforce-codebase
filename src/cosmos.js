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
      partitionKey: { paths: ["/partitionKey"] },
    });
    console.log("Cosmos DB initialized successfully");
    return container;
  } catch (error) {
    console.error("Error initializing Cosmos DB:", error.message);
    throw error;
  }
}

// Export initialized container
module.exports = {
  container: initializeCosmos(),
};