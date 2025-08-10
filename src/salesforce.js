const axios = require("axios");
const config = require("./config");

const SALESFORCE_INSTANCE_URL = "https://orgfarm-5a7d798f5f-dev-ed.develop.lightning.force.com";
const SALESFORCE_API_VERSION = "v60.0";

// Construct headers
function getHeaders() {
  return {
    Authorization: `Bearer ${config.salesforceAccessToken}`,
    "Content-Type": "application/json",
  };
}

// Create a new Salesforce lead
async function createSalesforceLead(context, state, leadData) {
  try {
    const response = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Lead`,
      {
        FirstName: leadData.firstName,
        LastName: leadData.lastName,
        Company: leadData.company,
      },
      {
        headers: getHeaders(),
      }
    );

    return {
      status: "success",
      id: response.data.id,
    };
  } catch (error) {
    console.error("Salesforce Create Lead Error:", error.response?.data || error.message);
    return {
      status: "error",
      message: error.response?.data?.[0]?.message || error.message,
    };
  }
}

// Get Salesforce leads (generic, supports parameterized limits)
async function getSalesforceLeads(limit = 20) {
  try {
    const query = `SELECT Id, FirstName, LastName, Company, Status, Email, Phone FROM Lead ORDER BY CreatedDate DESC LIMIT ${limit}`;
    const response = await axios.get(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/query?q=${encodeURIComponent(query)}`,
      {
        headers: getHeaders(),
      }
    );

    return {
      status: "success",
      records: response.data.records || [],
    };
  } catch (error) {
    console.error("Salesforce Get Leads Error:", error.response?.data || error.message);
    return {
      status: "error",
      message: error.response?.data?.[0]?.message || error.message,
    };
  }
}



// Update a lead by ID
async function updateSalesforceLead(context, state, leadId, fields) {
  try {
    const res = await axios.patch(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Lead/${leadId}`,
      fields,
      { headers: getHeaders() }
    );

    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Update Lead Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}


// Delete a Lead by company name
// Delete lead by ID
async function deleteSalesforceLead(context, state, leadId) {
  try {
    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Lead/${leadId}`,
      { headers: getHeaders() }
    );
    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Delete Lead Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}


module.exports = {
  createSalesforceLead,
  getSalesforceLeads,
  updateSalesforceLead,
  deleteSalesforceLead,
};
