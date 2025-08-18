require('dotenv').config({ path: './env/.env' });
require('dotenv').config({ path: './env/.env.local', override: true });

const config = {
  MicrosoftAppId: process.env.BOT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.BOT_TENANT_ID,
  MicrosoftAppPassword: process.env.BOT_PASSWORD,
  azureOpenAIKey: process.env.AZURE_OPENAI_API_KEY,
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
  azureOpenAIDeploymentName: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
  salesforceAccessToken: process.env.SALESFORCE_ACCESS_TOKEN,
  cosmosEndpoint: process.env.COSMOS_ENDPOINT,
  cosmosKey: process.env.COSMOS_KEY,
  cosmosDatabaseId: process.env.COSMOS_DATABASE_ID,
  cosmosContainerId: process.env.COSMOS_CONTAINER_ID,
  salesforceClientId: process.env.SALESFORCE_CLIENT_ID,
  salesforceClientSecret: process.env.SALESFORCE_CLIENT_SECRET,
  salesforceRedirectUri: process.env.SALESFORCE_REDIRECT_URI,
};
console.log("Config loaded:", config);
module.exports = config;
