const { getUserToken } = require("./cosmos");
const { createSalesforceClient } = require("./httpClient");

async function httpRequest(teamsChatId, url, method = "GET", data = null) {
  try {
    const { status, accessToken, instanceUrl } = await getUserToken(teamsChatId, "salesforce");
    
    if (!status) {
      throw new Error("User is not authenticated with Salesforce");
    }

    const client = createSalesforceClient(teamsChatId);
    const config = {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      }
    };

    let response;

    switch (method.toUpperCase()) {
      case "GET":
        console.log("Making GET request to:", `${instanceUrl}${url}`);
        response = await client.get(`${instanceUrl}${url}`, config);
        break;
      
      case "POST":
        response = await client.post(`${instanceUrl}${url}`, data, config);
        break;
      
      case "PUT":
        response = await client.put(`${instanceUrl}${url}`, data, config);
        break;

      case "PATCH":
        response = await client.patch(`${instanceUrl}${url}`, data, config);
        break;
      
      case "DELETE":
        response = await client.delete(`${instanceUrl}${url}`, config);
        break;
      
      default:
        throw new Error(`Unsupported HTTP method: ${method}`);
    }

    return response;
  } catch (error) {
    console.error(`Error making ${method} request:`, error);
    throw error;
  }
}



module.exports = { httpRequest };
