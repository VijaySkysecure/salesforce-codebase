const axios = require("axios");
const config = require("./config");
const {httpRequest} = require("./httpRequest");

const SALESFORCE_INSTANCE_URL = "https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0";
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
      `${SALESFORCE_INSTANCE_URL}/sobjects/Opportunity`,
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
// Helper function to find opportunity by name
async function findOpportunityByName(opportunityName) {
  try {
    const query = `SELECT Id, Name FROM Opportunity WHERE Name LIKE '%${opportunityName.replace(/'/g, "\\'")}%' LIMIT 10`;
    
    const res = await axios.get(
      `${SALESFORCE_INSTANCE_URL}/query?q=${encodeURIComponent(query)}`,
      { headers: getHeaders() }
    );

    return { 
      status: "success", 
      opportunities: res.data.records 
    };
  } catch (err) {
    console.error("Salesforce Find Opportunity Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

async function updateSalesforceOpportunity(context, state, opportunityId, fields) {
  try {
    const res = await axios.patch(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Opportunity/${opportunityId}`,
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

// Updated salesforce.js - Delete Function

async function deleteSalesforceOpportunity(context, state, opportunityIdentifier) {
  try {
    let opportunityId = opportunityIdentifier;
    let opportunityName = opportunityIdentifier;
    
    // If the identifier doesn't look like a Salesforce ID (15 or 18 chars starting with specific pattern)
    if (!opportunityIdentifier.match(/^[a-zA-Z0-9]{15}([a-zA-Z0-9]{3})?$/)) {
      // Try to find by name
      const searchResult = await findOpportunityByName(opportunityIdentifier);
      
      if (searchResult.status === "error") {
        return { status: "error", message: `Could not search for opportunity: ${searchResult.message}` };
      }
      
      if (searchResult.opportunities.length === 0) {
        return { status: "error", message: `No opportunity found with name containing: "${opportunityIdentifier}"` };
      }
      
      if (searchResult.opportunities.length > 1) {
        const names = searchResult.opportunities.map(opp => `"${opp.Name}"`).join(", ");
        return { 
          status: "error", 
          message: `Multiple opportunities found: ${names}. Please be more specific with the opportunity name.`,
          multipleResults: searchResult.opportunities
        };
      }
      
      opportunityId = searchResult.opportunities[0].Id;
      opportunityName = searchResult.opportunities[0].Name;
    }

    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Opportunity/${opportunityId}`,
      { headers: getHeaders() }
    );
    
    return { 
      status: "success", 
      opportunityId: opportunityId,
      opportunityName: opportunityName
    };
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
      `${SALESFORCE_INSTANCE_URL}/sobjects/Lead`,
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
      `${SALESFORCE_INSTANCE_URL}/sobjects/Lead/${leadId}`,
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
      `${SALESFORCE_INSTANCE_URL}/sobjects/Lead/${leadId}`,
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
      `${SALESFORCE_INSTANCE_URL}/sobjects/Task`,
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
      `${SALESFORCE_INSTANCE_URL}/sobjects/Task/${taskId}`,
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
async function deleteSalesforceTask(context, state, taskIdentifier) {
  try {
    let taskId = taskIdentifier;
    let taskSubject = taskIdentifier;

    // If the identifier doesn't look like a Salesforce ID
    if (!taskIdentifier.match(/^[a-zA-Z0-9]{15}([a-zA-Z0-9]{3})?$/)) {
      // Try to find by subject
      const searchResult = await findTaskBySubject(taskIdentifier);

      if (searchResult.status === "error") {
        return { status: "error", message: `Could not search for task: ${searchResult.message}` };
      }

      if (searchResult.tasks.length === 0) {
        return { status: "error", message: `No task found with subject containing: "${taskIdentifier}"` };
      }

      if (searchResult.tasks.length > 1) {
        const names = searchResult.tasks.map(t => `"${t.Subject}"`).join(", ");
        return { 
          status: "error", 
          message: `Multiple tasks found: ${names}. Please be more specific with the task subject.`,
          multipleResults: searchResult.tasks
        };
      }

      taskId = searchResult.tasks[0].Id;
      taskSubject = searchResult.tasks[0].Subject;
    }

    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Task/${taskId}`,
      { headers: getHeaders() }
    );

    return { 
      status: "success", 
      taskId: taskId,
      taskSubject: taskSubject
    };
  } catch (err) {
    console.error("Salesforce Delete Task Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}

async function findTaskBySubject(subjectFragment) {
  try {
    const query = `SELECT Id, Subject FROM Task WHERE Subject LIKE '%${subjectFragment.replace(/'/g, "\\'")}%' LIMIT 200`;

    const searchResponse = await axios.get(
      `${SALESFORCE_INSTANCE_URL}/query?q=${encodeURIComponent(query)}`,
      { headers: getHeaders() }
    );

    return { status: "success", tasks: searchResponse.data.records };
  } catch (err) {
    console.error("Salesforce Task Search Error:", err.response?.data || err.message);
    return { status: "error", message: err.message };
  }
}


// ======================= ACCOUNTS =======================

// Create a new Salesforce account
async function createSalesforceAccount(context, state, accountData) {
  try {
    const response = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Account`,
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
      `${SALESFORCE_INSTANCE_URL}/sobjects/Account/${accountId}`,
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
async function deleteSalesforceAccount(context, state, accountIdentifier) {
  try {
    let accountId = accountIdentifier;
    let accountName = accountIdentifier;

    // If the identifier doesn't look like a Salesforce ID
    if (!accountIdentifier.match(/^[a-zA-Z0-9]{15}([a-zA-Z0-9]{3})?$/)) {
      // Try to find by name
      const searchResult = await findAccountByName(accountIdentifier);

      if (searchResult.status === "error") {
        return { status: "error", message: `Could not search for account: ${searchResult.message}` };
      }

      if (searchResult.accounts.length === 0) {
        return { status: "error", message: `No account found with name containing: "${accountIdentifier}"` };
      }

      if (searchResult.accounts.length > 1) {
        const names = searchResult.accounts.map(acc => `"${acc.Name}"`).join(", ");
        return { 
          status: "error", 
          message: `Multiple accounts found: ${names}. Please be more specific with the account name.`,
          multipleResults: searchResult.accounts
        };
      }

      accountId = searchResult.accounts[0].Id;
      accountName = searchResult.accounts[0].Name;
    }

    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Account/${accountId}`,
      { headers: getHeaders() }
    );

    return { 
      status: "success", 
      accountId: accountId,
      accountName: accountName
    };
  } catch (err) {
    console.error("Salesforce Delete Account Error:", err.response?.data || err.message);
    return {
      status: "error",
      message: err.response?.data?.[0]?.message || err.message,
    };
  }
}
async function findAccountByName(nameFragment) {
  try {
    const query = `SELECT Id, Name FROM Account WHERE Name LIKE '%${nameFragment.replace(/'/g, "\\'")}%' LIMIT 200`;

    const searchResponse = await axios.get(
      `${SALESFORCE_INSTANCE_URL}/query?q=${encodeURIComponent(query)}`,
      { headers: getHeaders() }
    );

    return { status: "success", accounts: searchResponse.data.records };
  } catch (err) {
    console.error("Salesforce Account Search Error:", err.response?.data || err.message);
    return { status: "error", message: err.message };
  }
}



// ======================= CONTACTS =======================

// Create a new Salesforce contact
async function createSalesforceContact(context, state, contactData) {
  try {
    const response = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Contact`,
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

// Update a contact
async function updateSalesforceContact(context, state, contactId, fields) {
  try {
    const res = await axios.patch(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Contact/${contactId}`,
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

// Delete contact
async function deleteSalesforceContact(context, state, identifier) {
  try {
    let contactId = identifier;
    let contactName = null;

    // Check if identifier is a valid Salesforce Contact ID
    const isSalesforceId = /^003[0-9A-Za-z]{12}$/.test(identifier);

    if (!isSalesforceId) {
      console.log(`Looking up contact by name: ${identifier}`);

      const query = `SELECT Id, FirstName, LastName FROM Contact WHERE Name LIKE '%${identifier}%'`;
      const searchRes = await axios.get(
        `${SALESFORCE_INSTANCE_URL}/query?q=${encodeURIComponent(query)}`,
        { headers: getHeaders() }
      );

      const results = searchRes.data.records;

      if (results.length === 0) {
        return { status: "error", message: `No contact found with name "${identifier}"` };
      }
      if (results.length > 1) {
        return { status: "error", multipleResults: results };
      }

      contactId = results[0].Id;
      contactName = `${results[0].FirstName || ""} ${results[0].LastName || ""}`.trim();
    }

    // Delete by ID
    await axios.delete(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Contact/${contactId}`,
      { headers: getHeaders() }
    );

    return { status: "success", contactId, contactName };
  } catch (err) {
    console.error("Delete Contact Error:", err.response?.data || err.message);
    return { status: "error", message: err.response?.data?.[0]?.message || err.message };
  }
}



async function createSalesforceMeeting({ subject, startDateTime, endDateTime, whoId = null, whatId = null }) { 
  try {
    console.log("Creating meeting with times:", {
      receivedStart: startDateTime,
      receivedEnd: endDateTime,
      hasZSuffix: startDateTime.endsWith('Z'),
      parsedStart: new Date(startDateTime),
      parsedEnd: new Date(endDateTime)
    });

    const body = {
      Subject: subject,
      StartDateTime: startDateTime, // Should NOT have 'Z' suffix in Approach 2
      EndDateTime: endDateTime,     // Should NOT have 'Z' suffix in Approach 2
      IsAllDayEvent: false
    };

    if (whoId) body.WhoId = whoId;
    if (whatId) body.WhatId = whatId;

    console.log("Sending body to Salesforce:", JSON.stringify(body, null, 2));

    const res = await axios.post(
      `${SALESFORCE_INSTANCE_URL}/sobjects/Event`,
      body,
      { headers: getHeaders() }
    );

    console.log("Salesforce response:", res.data);
    return { status: "success", id: res.data.id };
  } catch (err) {
    console.error("Create Meeting Error:", err.response?.data || err.message);
    return { status: "error", message: err.response?.data?.[0]?.message || err.message };
  }
}



/**
 * Utility: Search Contact or Account by name
 */
async function findSalesforceContactOrAccount(name) {
  const contactQuery = `SELECT Id, FirstName, LastName FROM Contact WHERE Name LIKE '%${name}%' LIMIT 1`;
  const accountQuery = `SELECT Id, Name FROM Account WHERE Name LIKE '%${name}%' LIMIT 1`;

  const contactRes = await axios.get(
    `${SALESFORCE_INSTANCE_URL}/query?q=${encodeURIComponent(contactQuery)}`,
    { headers: getHeaders() }
  );

  if (contactRes.data.records.length > 0) {
    return { type: "Contact", ...contactRes.data.records[0] };
  }

  const accountRes = await axios.get(
    `${SALESFORCE_INSTANCE_URL}/query?q=${encodeURIComponent(accountQuery)}`,
    { headers: getHeaders() }
  );

  if (accountRes.data.records.length > 0) {
    return { type: "Account", ...accountRes.data.records[0] };
  }

  return null;
}


// NEW FUNCTION: Find Meeting by Subject or DateTime
async function findSalesforceMeeting(data, teamsChatId) {
  try {
    let query;
    
    if (data.subject) {
      // Search by subject (case-insensitive, partial match)
      query = `SELECT Id, Subject, StartDateTime, EndDateTime, WhoId, WhatId FROM Event WHERE Subject LIKE '%${data.subject}%' ORDER BY StartDateTime DESC LIMIT 1`;
    } else if (data.dateTime) {
      // Convert the provided datetime to a range (±30 minutes for flexibility)
      const searchMoment = moment(data.dateTime);
      const startRange = searchMoment.clone().subtract(30, 'minutes').format('YYYY-MM-DDTHH:mm:ss.SSSZ');
      const endRange = searchMoment.clone().add(30, 'minutes').format('YYYY-MM-DDTHH:mm:ss.SSSZ');
      
      query = `SELECT Id, Subject, StartDateTime, EndDateTime, WhoId, WhatId FROM Event WHERE StartDateTime >= ${startRange} AND StartDateTime <= ${endRange} ORDER BY StartDateTime ASC LIMIT 1`;
    } else {
      return null;
    }

    console.log("Meeting search query:", query);
    console.log("Teams Chat ID for request:", teamsChatId);

    const response = await httpRequest(
      teamsChatId,
      `/services/data/v61.0/query?q=${encodeURIComponent(query)}`,
      "GET"
    );

    console.log("Meeting search results:", response.data);

    if (response.data.records.length > 0) {
      return response.data.records[0];
    }

    return null;
  } catch (err) {
    console.error("Find Meeting Error:", err);
    return null;
  }
}

// NEW FUNCTION: Update Meeting
async function updateSalesforceMeeting(teamsChatId, meetingId, { startDateTime, endDateTime }) {
  try {
    console.log("Updating meeting with ID:", meetingId);
    console.log("New times:", {
      startDateTime,
      endDateTime
    });

    const body = {
      StartDateTime: startDateTime,
      EndDateTime: endDateTime
    };

    console.log("Sending update body to Salesforce:", JSON.stringify(body, null, 2));

    const response = await httpRequest(
      teamsChatId,
      `/services/data/v61.0/sobjects/Event/${meetingId}`,
      "PATCH",
      body
    );

    console.log("Salesforce update response:", response.status);
    return { status: "success" };
  } catch (err) {
    console.error("Update Meeting Error:", err.response?.data || err.message);
    return { status: "error", message: err.response?.data?.[0]?.message || err.message };
  }
}

// NEW FUNCTION: Cancel Meeting
async function cancelSalesforceMeeting(teamsChatId, meetingId) {
  try {
    console.log("Cancelling meeting with ID:", meetingId);

    const response = await httpRequest(
      teamsChatId,
      `/services/data/v61.0/sobjects/Event/${meetingId}`,
      "DELETE"
    );

    console.log("Salesforce cancel response:", response.status);
    return { status: "success" };
  } catch (err) {
    console.error("Cancel Meeting Error:", err.response?.data || err.message);
    return { status: "error", message: err.response?.data?.[0]?.message || err.message };
  }
}


async function generateOpportunityFromEmail(context, state, email, teamsChatId) {
  try {
    console.log("Processing email for opportunity creation:", {
      subject: email.subject,
      from: email.from.emailAddress.name,
      fromAddress: email.from.emailAddress.address
    });

    // Generate opportunity name from email subject and sender (2-3 words)
    let opportunityName = "";
    const senderName = email.from.emailAddress.name || email.from.emailAddress.address.split('@')[0];
    const subject = email.subject || "Email";
    
    // Extract key words from subject and combine with sender name
    const subjectWords = subject.split(' ').filter(word => 
      word.length > 3 && 
      !['the', 'and', 'for', 'with', 'from', 'that', 'this', 'your'].includes(word.toLowerCase())
    );
    
    if (subjectWords.length > 0) {
      // Take first meaningful word from subject + sender name
      const keyWord = subjectWords[0];
      opportunityName = `${senderName} ${keyWord}`.substring(0, 50); // Limit length
    } else {
      // Fallback to sender name + "Opportunity"
      opportunityName = `${senderName} Opportunity`.substring(0, 50);
    }

    // Set close date to 31 days from today
    const closeDate = new Date();
    closeDate.setDate(closeDate.getDate() + 31);
    const formattedCloseDate = closeDate.toISOString().split('T')[0]; // YYYY-MM-DD format

    const stageName = "Prospecting";

    console.log("Creating Salesforce opportunity with details:", {
      name: opportunityName,
      stageName: stageName,
      closeDate: formattedCloseDate
    });

    // Create the opportunity in Salesforce
    const response = await httpRequest(
      teamsChatId, 
      `/services/data/v59.0/sobjects/Opportunity`, 
      "POST", 
      { 
        Name: opportunityName, 
        StageName: stageName, 
        CloseDate: formattedCloseDate 
      }
    );

    if (response.data.id) {
      return {
        success: true,
        id: response.data.id,
        name: opportunityName,
        stageName: stageName,
        closeDate: formattedCloseDate
      };
    } else {
      return {
        success: false,
        message: response.message || "Unknown error occurred while creating opportunity"
      };
    }

  } catch (error) {
    console.error("Error in generateOpportunityFromEmail:", error);
    return {
      success: false,
      message: error.response?.data?.error?.message || error.message || "Failed to create opportunity"
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
  createSalesforceMeeting,
  findSalesforceContactOrAccount,
  findSalesforceMeeting,
  updateSalesforceMeeting,
  cancelSalesforceMeeting,
  generateOpportunityFromEmail
};