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

// ======================= OPPORTUNITIES =======================

// Create a new Salesforce opportunity
async function createSalesforceOpportunity(context, state, opportunityData) {
  try {
    const response = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Opportunity`,
      {
        Name: opportunityData.name,
        StageName: opportunityData.stageName,
        CloseDate: opportunityData.closeDate,
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
    console.error("Salesforce Create Opportunity Error:", error.response?.data || error.message);
    return {
      status: "error",
      message: error.response?.data?.[0]?.message || error.message,
    };
  }
}

// Update an opportunity by ID
async function updateSalesforceOpportunity(context, state, opportunityId, fields) {
  try {
    const res = await axios.patch(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Opportunity/${opportunityId}`,
      fields,
      { headers: getHeaders() }
    );

    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Update Opportunity Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

// Delete opportunity by ID
async function deleteSalesforceOpportunity(context, state, opportunityId) {
  try {
    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Opportunity/${opportunityId}`,
      { headers: getHeaders() }
    );
    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Delete Opportunity Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

// ======================= LEADS =======================

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

// ======================= TASKS =======================

// Create a new Salesforce task
async function createSalesforceTask(context, state, taskData) {
  try {
    const response = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Task`,
      {
        Subject: taskData.subject,
        Status: taskData.status,
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
    console.error("Salesforce Create Task Error:", error.response?.data || error.message);
    return {
      status: "error",
      message: error.response?.data?.[0]?.message || error.message,
    };
  }
}

// Update a task by ID
async function updateSalesforceTask(context, state, taskId, fields) {
  try {
    const res = await axios.patch(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Task/${taskId}`,
      fields,
      { headers: getHeaders() }
    );

    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Update Task Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

// Delete task by ID
async function deleteSalesforceTask(context, state, taskId) {
  try {
    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Task/${taskId}`,
      { headers: getHeaders() }
    );
    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Delete Task Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

// ======================= ACCOUNTS =======================

// Create a new Salesforce account
async function createSalesforceAccount(context, state, accountData) {
  try {
    const response = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Account`,
      {
        Name: accountData.name,
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
    console.error("Salesforce Create Account Error:", error.response?.data || error.message);
    return {
      status: "error",
      message: error.response?.data?.[0]?.message || error.message,
    };
  }
}

// Update an account by ID
async function updateSalesforceAccount(context, state, accountId, fields) {
  try {
    const res = await axios.patch(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Account/${accountId}`,
      fields,
      { headers: getHeaders() }
    );

    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Update Account Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

// Delete account by ID
async function deleteSalesforceAccount(context, state, accountId) {
  try {
    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Account/${accountId}`,
      { headers: getHeaders() }
    );
    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Delete Account Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

// ======================= CONTACTS =======================

// Create a new Salesforce contact
async function createSalesforceContact(context, state, contactData) {
  try {
    const response = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Contact`,
      {
        FirstName: contactData.firstName,
        LastName: contactData.lastName,
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
    console.error("Salesforce Create Contact Error:", error.response?.data || error.message);
    return {
      status: "error",
      message: error.response?.data?.[0]?.message || error.message,
    };
  }
}

// Update a contact by ID
async function updateSalesforceContact(context, state, contactId, fields) {
  try {
    const res = await axios.patch(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Contact/${contactId}`,
      fields,
      { headers: getHeaders() }
    );

    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Update Contact Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

// Delete contact by ID
async function deleteSalesforceContact(context, state, contactId) {
  try {
    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/services/data/${SALESFORCE_API_VERSION}/sobjects/Contact/${contactId}`,
      { headers: getHeaders() }
    );
    return { status: "success" };
  } catch (err) {
    console.error("Salesforce Delete Contact Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

module.exports = {
  // Opportunities
  createSalesforceOpportunity,
  updateSalesforceOpportunity,
  deleteSalesforceOpportunity,
  // Leads
  createSalesforceLead,
  getSalesforceLeads,
  updateSalesforceLead,
  deleteSalesforceLead,
  // Tasks
  createSalesforceTask,
  updateSalesforceTask,
  deleteSalesforceTask,
  // Accounts
  createSalesforceAccount,
  updateSalesforceAccount,
  deleteSalesforceAccount,
  // Contacts
  createSalesforceContact,
  updateSalesforceContact,
  deleteSalesforceContact,
};