const { getUserToken } = require("./cosmos");
const { createSalesforceClient, createOutlookClient } = require("./httpClient");

// Common function to handles the http request and response related things.
async function httpResponse(client, method = "GET", url, data, config) {
  try {
    let response;

    switch (method.toUpperCase()) {
      case "GET":
        console.log("Making GET request to:", `${url}`);
        response = await client.get(url, config);
        break;

      case "POST":
        response = await client.post(url, data, config);
        break;

      case "PUT":
        response = await client.put(url, data, config);
        break;

      case "PATCH":
        response = await client.patch(url, data, config);
        break;

      case "DELETE":
        response = await client.delete(url, config);
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
    response = await httpResponse(client, method, `${instanceUrl}${url}`, data, config)

    return response;
  } catch (error) {
    console.error(`Error making ${method} request:`, error);
    throw error;
  }
}

async function outlookHttpRequest(teamsChatId, url, method = "GET", data = null) {
  try {
    const { status, accessToken } = await getUserToken(teamsChatId, "outlook");

    if (!status) {
      throw new Error("User is not authenticated with Salesforce");
    }

    const client = createOutlookClient(teamsChatId);
    const config = {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      }
    };

    let response;
    response = await httpResponse(client, method, url, data, config)

    return response;
  } catch (error) {
    console.error(`Error making ${method} request:`, error);
    throw error;
  }
}





module.exports = { httpRequest, outlookHttpRequest };
