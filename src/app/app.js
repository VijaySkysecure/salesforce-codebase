const { MemoryStorage, MessageFactory, CardFactory, ActivityTypes } = require("botbuilder");
const path = require("path");
const config = require("../config");

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
const { Application, ActionPlanner, OpenAIModel, PromptManager } = require("@microsoft/teams-ai");

// Create AI components
const model = new OpenAIModel({
  azureApiKey: config.azureOpenAIKey,
  azureDefaultDeployment: config.azureOpenAIDeploymentName,
  azureEndpoint: config.azureOpenAIEndpoint,

  useSystemMessages: true,
  logRequests: true,
});
const prompts = new PromptManager({
  promptsFolder: path.join(__dirname, "../prompts"),
});
const planner = new ActionPlanner({
  model,
  prompts,
  defaultPrompt: "chat",
});

// Define storage and application
const storage = new MemoryStorage();
const app = new Application({
  storage,
  ai: {
    planner,
    // enable_feedback_loop: true,
  },
});

// app.feedbackLoop(async (context, state, feedbackLoopData) => {
//   //add custom feedback process logic here
//   console.log("Your feedback is " + JSON.stringify(context.activity.value));
// });



const {
  createSalesforceLead,
  getSalesforceLeads,
  updateSalesforceLead,
  deleteSalesforceLead,
    // Opportunities
  createSalesforceOpportunity,
  updateSalesforceOpportunity,
  deleteSalesforceOpportunity,
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
} = require("../salesforce");


const defaultConversationState = {
  isAuthenticated: false,
  userId: null,
  
  // Last fetched raw data
  lastAccountsData: null,
  lastLeadsData: null,
  lastDealsData: null,
  lastContactsData: null,
  lastTasksData: null,
  lastMeetingsData: null,
  lastCampaignsData: null,
  lastCallsData: null,

  // Counts for each record type
  accountCount: 0,
  leadsCount: 0,
  dealsCount: 0,
  contactsCount: 0,
  tasksCount: 0,
  meetingsCount: 0,
  campaignsCount: 0,
  callsCount: 0,

  // Display-friendly JSON strings
  formattedAccounts: null,
  formattedLeads: null,
  formattedDeals: null,
  formattedContacts: null,
  formattedTasks: null,
  formattedMeetings: null,
  formattedCampaigns: null,
  formattedCalls: null,
  formattedOpportunities: null,
  
  // Search results
  formattedLeadsSearch: null,

  // Raw API response storage
  lastRawAccountsResponse: null,

  // User-specific
  callReminders: [],

  // Lead-specific request metadata
  requestedLeadLimit: null
};



// Initialize conversation state helper
function initializeConversationState(state) {
  if (!state.conversation) {
    state.conversation = { ...defaultConversationState };
  }
  Object.keys(defaultConversationState).forEach((key) => {
    if (state.conversation[key] === undefined) {
      state.conversation[key] = defaultConversationState[key];
    }
  });
}


app.ai.action("GetSalesforceLeads", async (context, state, parameters) => {
  try {
    console.log("GetSalesforceLeads action called with parameters:", parameters);
    // Ensure state is initialized with defaultConversationState
    if (!state.conversation) state.conversation = {};
    state.conversation = { ...defaultConversationState, ...state.conversation };
    

    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    // Set and store limit in state (like Zoho function does)
    const limit = Math.min(parameters.limit || 20, 200);
    state.conversation.requestedLeadLimit = limit;

    // Salesforce query
    const query = `SELECT Id, FirstName, LastName, Company, Status, Email, Phone 
                   FROM Lead ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const config = require("../config");
    const axios = require("axios");
    console.log("djwffeiu",
    {
      Authorization: `Bearer ${config.salesforceAccessToken}`,
      "Content-Type": "application/json",
    }
    )
    const headers = {
      Authorization: `Bearer ${config.salesforceAccessToken}`,
      "Content-Type": "application/json",
    };

    const response = await axios.get(
      `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(query)}`,
      { headers }
    );

    const records = response.data.records || [];
    console.log(response.data, "egbj");

    if (records.length === 0) {
      await context.sendActivity(
        MessageFactory.text("📊 No leads found in your Salesforce CRM.")
      );
      return "No leads found";
    }

    // Store raw data
    state.conversation.lastLeadsData = records;
    state.conversation.leadsCount = records.length;

    // Format for display
    const formattedLeads = records.map((l) => ({
      id: l.Id,
      name: `${l.FirstName || ''} ${l.LastName || ''}`.trim() || "—",
      company: l.Company || "—",
      status: l.Status || "—",
      email: l.Email || "—",
      phone: l.Phone || "—",
    }));

    // Store formatted version
    // state.conversation.formattedLeads = JSON.stringify(formattedLeads, null, 2);
 

const adaptiveCard = {
  type: "AdaptiveCard",
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  version: "1.4",
  body: [
    {
      type: "TextBlock",
      text: `📊 Retrieved ${records.length} Leads`,
      weight: "Bolder",
      size: "Large",
      wrap: true
    },
    ...formattedLeads.map(lead => ({
      type: "Container",
      items: [
        { type: "TextBlock", text: `**Name:** ${lead.name}`, wrap: true },
        { type: "TextBlock", text: `**Company:** ${lead.company}`, wrap: true },
        { type: "TextBlock", text: `**Status:** ${lead.status}`, wrap: true },
        { type: "TextBlock", text: `**Email:** ${lead.email}`, wrap: true },
        { type: "TextBlock", text: `**Phone:** ${lead.phone}`, wrap: true }
      ],
      separator: true
    }))
  ]
};

await context.sendActivity({
  attachments: [CardFactory.adaptiveCard(adaptiveCard)]
});

    // return response.data.records
    return `Retrieved ${records.length} leads successfully`;
  } catch (error) {
    console.error("Error fetching Salesforce leads:", error);
    const errorMessage = error.response?.data?.[0]?.message || error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error retrieving leads: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});



app.ai.action("CreateSalesforceLead", async (context, state, parameters) => {
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const firstName = parameters.firstName;
    const lastName = parameters.lastName || "";
    const company = parameters.company || "";

    if (!firstName || !lastName || !company) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `firstName`, `lastName`, and `company`.")
      );
      return "Missing required parameters";
    }

    const response = await createSalesforceLead(context, state, { firstName, lastName, company });

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Lead Created Successfully!**\n\n` +
          `🆔 **Lead ID:** ${response.id}\n` +
          `👤 **Name:** ${firstName} ${lastName}\n` +
          `🏢 **Company:** ${company}`
        )
      );
      return `Successfully created Salesforce lead ${firstName} ${lastName}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to create lead: ${response.message || "Unknown error"}.`)
      );
      return `Failed to create lead: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error creating Salesforce lead:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error creating lead: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});



app.ai.action("UpdateSalesforceLead", async (context, state, parameters) => {
  try {
    console.log("UpdateSalesforceLead called with:", parameters);
    initializeConversationState(state);
 
    const config = require("../config");
    const axios = require("axios");
 
    let leadId = parameters.leadId;
 
    // 1. Try finding lead by email first
    if (!leadId && parameters.email) {
      const emailQuery = `SELECT Id, FirstName, LastName FROM Lead WHERE Email LIKE '%${parameters.email}%' LIMIT 200`;
      const emailResponse = await axios.get(
        `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(emailQuery)}`,
        { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      );
 
      if (emailResponse.data.records.length === 1) {
        leadId = emailResponse.data.records[0].Id;
      } else if (emailResponse.data.records.length > 1) {
        await context.sendActivity("⚠ Multiple leads found by email. Please refine search.");
        return;
      }
    }
 
    // 2. If still no leadId, try by name
    if (!leadId && (parameters.firstName || parameters.lastName || parameters.name)) {
      let nameQuery;
      if (parameters.name) {
        nameQuery = `SELECT Id, FirstName, LastName FROM Lead WHERE Name LIKE '%${parameters.name}%' LIMIT 200`;
      } else {
        const firstNamePart = parameters.firstName ? `FirstName LIKE '%${parameters.firstName}%'` : "";
        const lastNamePart = parameters.lastName ? `LastName LIKE '%${parameters.lastName}%'` : "";
        const andClause = firstNamePart && lastNamePart ? " AND " : "";
        nameQuery = `SELECT Id, FirstName, LastName FROM Lead WHERE ${firstNamePart}${andClause}${lastNamePart} LIMIT 200`;
      }
 
      const nameResponse = await axios.get(
        `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
        { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      );
 
      if (nameResponse.data.records.length === 0) {
        await context.sendActivity("❌ No matching Salesforce lead found.");
        return;
      }
      if (nameResponse.data.records.length > 1) {
        await context.sendActivity("⚠ Multiple leads found by name. Please refine search.");
        return;
      }
      leadId = nameResponse.data.records[0].Id;
    }
 
    if (!leadId) {
      await context.sendActivity("❌ Could not find a matching lead by provided details.");
      return;
    }
 
    // 3. Build update object
    const updateFields = {
      ...(parameters.firstName && { FirstName: parameters.firstName }),
      ...(parameters.lastName && { LastName: parameters.lastName }),
      ...(parameters.company && { Company: parameters.company }),
      ...(parameters.email && { Email: parameters.email }),
      ...(parameters.phone && { Phone: parameters.phone }),
      ...(parameters.status && { Status: parameters.status }),
      ...(parameters.title && { Title: parameters.title }),
      ...(parameters.leadSource && { LeadSource: parameters.leadSource }),
      ...(parameters.industry && { Industry: parameters.industry })
    };
 
    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity("❌ No fields provided to update.");
      return;
    }
 
    console.log(`Updating lead ${leadId} with fields:`, updateFields);
    const { updateSalesforceLead } = require("../salesforce");
    const response = await updateSalesforceLead(context, state, leadId, updateFields);
 
    if (response.status === "success") {
      await context.sendActivity(`✅ Lead updated successfully (ID: ${leadId}).`);
    } else {
      await context.sendActivity(`❌ Failed to update lead: ${response.message || "Unknown error"}.`);
    }
  } catch (error) {
    console.error("Error in UpdateSalesforceLead:", error);
    await context.sendActivity(`❌ Error: ${error.message || "Unknown error"}`);
  }
});




app.ai.action("DeleteSalesforceLead", async (context, state, parameters) => {
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const config = require("../config");
    const axios = require("axios");

    // 1. If no leadId is given, try to find it by name
    let leadId = parameters.leadId;
    if ((parameters.firstName || parameters.lastName || parameters.name)) {
      let nameQuery;
      if (parameters.name) {
        // Full name search
        nameQuery = `SELECT Id, FirstName, LastName FROM Lead WHERE Name LIKE '%${parameters.name}%' LIMIT 200`;
      } else {
        // Partial name search
        const firstNamePart = parameters.firstName ? `FirstName LIKE '%${parameters.firstName}%'` : "";
        const lastNamePart = parameters.lastName ? `LastName LIKE '%${parameters.lastName}%'` : "";
        const andClause = firstNamePart && lastNamePart ? " AND " : "";
        nameQuery = `SELECT Id, FirstName, LastName FROM Lead WHERE ${firstNamePart}${andClause}${lastNamePart} LIMIT 200`;
      }

      const headers = {
        Authorization: `Bearer ${config.salesforceAccessToken}`,
        "Content-Type": "application/json",
      };

      const queryUrl = `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`;
      const lookupResponse = await axios.get(queryUrl, { headers });

      if (lookupResponse.data.records.length === 0) {
        await context.sendActivity(
          MessageFactory.text(`❌ No Salesforce lead found matching the provided name.`)
        );
        return "No lead found by name";
      }
      if (lookupResponse.data.records.length > 1) {
        await context.sendActivity(
          MessageFactory.text(`⚠ Multiple leads found matching the provided name. Please refine your search.`)
        );
        return "Multiple leads found by name";
      }

      leadId = lookupResponse.data.records[0].Id;
    }

    if (!leadId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need either: Lead ID or a name to find the lead.")
      );
      return "Missing required parameters";
    }

    console.log(`Deleting lead ${leadId} from Salesforce CRM...`);
    const { deleteSalesforceLead } = require("../salesforce");

    const response = await deleteSalesforceLead(context, state, leadId);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(`✅ **Lead Deleted Successfully!**\n\n🆔 **Lead ID:** ${leadId}`)
      );
      return `Successfully deleted lead ${leadId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to delete lead: ${response.message || "Unknown error"}.`)
      );
      return `Failed to delete lead: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error deleting Salesforce lead:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error deleting lead: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});



// ======================= OPPORTUNITIES =======================

app.ai.action("GetSalesforceOpportunities", async (context, state, parameters) => {
  console.log("GetSalesforceOpportunities action called with parameters:", parameters);
  try {
    initializeConversationState(state);
    
    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, Name, StageName, Amount, CloseDate, AccountId, Account.Name FROM Opportunity ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const axios = require("axios");
    
    const headers = {
      Authorization: `Bearer ${config.salesforceAccessToken}`,
      "Content-Type": "application/json",
    };

    const response = await axios.get(
      `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(query)}`,
      { headers }
    );

    const records = response.data.records || [];

    if (records.length === 0) {
      await context.sendActivity(MessageFactory.text("📊 No opportunities found in your Salesforce CRM."));
      return "No opportunities found";
    }

    const formattedOpportunities = records.map((o) => ({
      id: o.Id,
      name: o.Name || "—",
      stage: o.StageName || "—",
      amount: o.Amount ? `$${o.Amount.toLocaleString()}` : "—",
      closeDate: o.CloseDate || "—",
      accountName: o.Account?.Name || "—",
    }));

    state.conversation.formattedOpportunities = JSON.stringify(formattedOpportunities, null, 2);

    // Create Adaptive Card
    const adaptiveCard = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "🎯 Your Recent Opportunities",
          size: "Large",
          weight: "Bolder",
          color: "Light"
        },
        {
          type: "TextBlock",
          text: `Found ${records.length} opportunities`,
          size: "Medium",
          color: "Good",
          spacing: "Small"
        },
        ...formattedOpportunities.slice(0, 10).map((opportunity, index) => ({
          type: "Container",
          style: "emphasis",
          spacing: "Medium",
          items: [
            {
              type: "ColumnSet",
              columns: [
                {
                  type: "Column",
                  width: "stretch",
                  items: [
                    {
                      type: "TextBlock",
                      text: `${index + 1}. ${opportunity.name}`,
                      weight: "Bolder",
                      size: "Medium",
                      wrap: true,
                      color: "Light"
                    },
                    {
                      type: "TextBlock",
                      text: `Account: ${opportunity.accountName}`,
                      color: "Light",
                      size: "Small",
                      spacing: "None"
                    }
                  ]
                },
                {
                  type: "Column",
                  width: "auto",
                  items: [
                    {
                      type: "TextBlock",
                      text: opportunity.amount,
                      weight: "Bolder",
                      size: "Medium",
                      color: "Good",
                      horizontalAlignment: "Right"
                    }
                  ]
                }
              ]
            },
            {
              type: "ColumnSet",
              spacing: "Small",
              columns: [
                {
                  type: "Column",
                  width: "stretch",
                  items: [
                    {
                      type: "TextBlock",
                      text: `Stage: ${opportunity.stage}`,
                      size: "Small",
                      color: "Light"
                    }
                  ]
                },
                {
                  type: "Column",
                  width: "auto",
                  items: [
                    {
                      type: "TextBlock",
                      text: `Close: ${opportunity.closeDate}`,
                      size: "Small",
                      color: "Light",
                      horizontalAlignment: "Right"
                    }
                  ]
                }
              ]
            }
          ]
        }))
      ]
    };

    // If there are more than 10 opportunities, add a note
    if (formattedOpportunities.length > 10) {
      adaptiveCard.body.push({
        type: "TextBlock",
        text: `... and ${formattedOpportunities.length - 10} more opportunities`,
        size: "Small",
        color: "Attention",
        horizontalAlignment: "Center",
        spacing: "Medium"
      });
    }

    const cardAttachment = MessageFactory.attachment({
      contentType: "application/vnd.microsoft.card.adaptive",
      content: adaptiveCard
    });

    await context.sendActivity(cardAttachment);
    
    return `Retrieved ${records.length} opportunities successfully`;
  } catch (error) {
    console.error("Error fetching Salesforce opportunities:", error);
    const errorMessage = error.response?.data?.[0]?.message || error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error retrieving opportunities: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("CreateSalesforceOpportunity", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const name = parameters.name;
    const stageName = parameters.stageName || "Prospecting";
    const closeDate = parameters.closeDate;

    if (!name || !closeDate) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `name` and `closeDate`.")
      );
      return "Missing required parameters";
    }

    const response = await createSalesforceOpportunity(context, state, { name, stageName, closeDate });

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Opportunity Created Successfully!**\n\n` +
          `🆔 **Opportunity ID:** ${response.id}\n` +
          `📋 **Name:** ${name}\n` +
          `📊 **Stage:** ${stageName}\n` +
          `📅 **Close Date:** ${closeDate}`
        )
      );
      return `Successfully created Salesforce opportunity ${name}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to create opportunity: ${response.message || "Unknown error"}.`)
      );
      return `Failed to create opportunity: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error creating Salesforce opportunity:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error creating opportunity: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});



app.ai.action("UpdateSalesforceOpportunity", async (context, state, parameters) => {
  try {
    console.log("UpdateSalesforceOpportunity called with:", parameters);
    initializeConversationState(state);
 
    const config = require("../config");
    const axios = require("axios");
 
    let opportunityId = parameters.opportunityId;
 
    // 1. Try finding opportunity by name first if no ID provided
    if (!opportunityId && (parameters.opportunityName || parameters.name)) {
      const opportunityName = parameters.opportunityName || parameters.name;
      const nameQuery = `SELECT Id, Name FROM Opportunity WHERE Name LIKE '%${opportunityName}%' LIMIT 200`;
      
      const nameResponse = await axios.get(
        `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
        { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      );
 
      if (nameResponse.data.records.length === 1) {
        opportunityId = nameResponse.data.records[0].Id;
      } else if (nameResponse.data.records.length > 1) {
        await context.sendActivity("⚠ Multiple opportunities found by name. Please refine search.");
        return;
      } else if (nameResponse.data.records.length === 0) {
        await context.sendActivity("❌ No matching Salesforce opportunity found.");
        return;
      }
    }
 
    if (!opportunityId) {
      await context.sendActivity("❌ Could not find a matching opportunity by provided details.");
      return;
    }
 
    // 2. Build update object
    const updateFields = {
      ...(parameters.newName && { Name: parameters.newName }),
      ...(parameters.stageName && { StageName: parameters.stageName }),
      ...(parameters.stage && { StageName: parameters.stage }),
      ...(parameters.amount && { Amount: parseFloat(parameters.amount.toString().replace(/[$,]/g, '')) }),
      ...(parameters.closeDate && { CloseDate: parameters.closeDate }),
      ...(parameters.description && { Description: parameters.description }),
      ...(parameters.probability && { Probability: parseInt(parameters.probability) }),
      ...(parameters.accountId && { AccountId: parameters.accountId })
    };
 
    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity("❌ No fields provided to update.");
      return;
    }
 
    console.log(`Updating opportunity ${opportunityId} with fields:`, updateFields);
    const { updateSalesforceOpportunity } = require("../salesforce");
    const response = await updateSalesforceOpportunity(context, state, opportunityId, updateFields);
 
    if (response.status === "success") {
      await context.sendActivity(`✅ Opportunity updated successfully (ID: ${opportunityId}).`);
    } else {
      await context.sendActivity(`❌ Failed to update opportunity: ${response.message || "Unknown error"}.`);
    }
  } catch (error) {
    console.error("Error in UpdateSalesforceOpportunity:", error);
    await context.sendActivity(`❌ Error: ${error.message || "Unknown error"}`);
  }
});

// Updated app.js - Delete Action

app.ai.action("DeleteSalesforceOpportunity", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    // Accept both opportunityId and opportunityName parameters
    const opportunityIdentifier = parameters.opportunityId || parameters.opportunityName || parameters.name;

    if (!opportunityIdentifier) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need either the Opportunity ID or Opportunity Name to delete an opportunity.")
      );
      return "Missing required parameters";
    }

    console.log(`Deleting opportunity "${opportunityIdentifier}" from Salesforce CRM...`);

    const response = await deleteSalesforceOpportunity(context, state, opportunityIdentifier);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Opportunity Deleted Successfully!**\n\n` +
          `🆔 **Opportunity ID:** ${response.opportunityId}\n` +
          `📋 **Opportunity Name:** ${response.opportunityName}`
        )
      );
      return `Successfully deleted Salesforce opportunity "${response.opportunityName}" (${response.opportunityId})`;
    } else if (response.multipleResults) {
      // Handle multiple results case
      const opportunityList = response.multipleResults
        .map((opp, index) => `${index + 1}. ${opp.Name} (${opp.Id})`)
        .join('\n');
      
      await context.sendActivity(
        MessageFactory.text(
          `❌ **Multiple opportunities found:**\n\n${opportunityList}\n\n` +
          `Please be more specific with the opportunity name, or use the exact Opportunity ID.`
        )
      );
      return `Multiple opportunities found with similar names`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to delete opportunity: ${response.message || "Unknown error"}.`)
      );
      return `Failed to delete opportunity: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error deleting Salesforce opportunity:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error deleting opportunity: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

// ======================= TASKS =======================

app.ai.action("GetSalesforceTasks", async (context, state, parameters) => {
  console.log("GetSalesforceTasks action called with parameters:", parameters);
  try {
    initializeConversationState(state);
    
    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, Subject, Status, Priority, ActivityDate, Description FROM Task ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const axios = require("axios");
    
    const headers = {
      Authorization: `Bearer ${config.salesforceAccessToken}`,
      "Content-Type": "application/json",
    };

    const response = await axios.get(
      `https://orgfarm-5a7d798f5f-dev-ed.develop.lightning.force.com/services/data/v60.0/query?q=${encodeURIComponent(query)}`,
      { headers }
    );

    const records = response.data.records || [];

    if (records.length === 0) {
      await context.sendActivity(MessageFactory.text("📊 No tasks found in your Salesforce CRM."));
      return "No tasks found";
    }

    const formattedTasks = records.map((t) => ({
      id: t.Id,
      subject: t.Subject || "—",
      status: t.Status || "—",
      priority: t.Priority || "—",
      activityDate: t.ActivityDate || "—",
      description: t.Description || "—",
    }));

    state.conversation.formattedTasks = JSON.stringify(formattedTasks, null, 2);
    return `Retrieved ${records.length} tasks successfully`;
  } catch (error) {
    console.error("Error fetching Salesforce tasks:", error);
    const errorMessage = error.response?.data?.[0]?.message || error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error retrieving tasks: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("CreateSalesforceTask", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const subject = parameters.subject;
    const status = parameters.status || "Not Started";

    if (!subject) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `subject`.")
      );
      return "Missing required parameters";
    }

    const response = await createSalesforceTask(context, state, { subject, status });

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Task Created Successfully!**\n\n` +
          `🆔 **Task ID:** ${response.id}\n` +
          `📋 **Subject:** ${subject}\n` +
          `📊 **Status:** ${status}`
        )
      );
      return `Successfully created Salesforce task ${subject}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to create task: ${response.message || "Unknown error"}.`)
      );
      return `Failed to create task: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error creating Salesforce task:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error creating task: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("UpdateSalesforceTask", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    if (!parameters.taskId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Task ID to update a task.")
      );
      return "Missing required parameters";
    }

    const updateFields = {
      ...(parameters.subject && { Subject: parameters.subject }),
      ...(parameters.status && { Status: parameters.status }),
      ...(parameters.priority && { Priority: parameters.priority }),
      ...(parameters.activityDate && { ActivityDate: parameters.activityDate }),
      ...(parameters.description && { Description: parameters.description }),
    };

    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity(
        MessageFactory.text("❌ No fields provided to update. Please include at least one field like subject, status, etc.")
      );
      return "No fields provided to update";
    }

    console.log(`Updating task ${parameters.taskId} in Salesforce CRM...`);

    const response = await updateSalesforceTask(context, state, parameters.taskId, updateFields);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Task Updated Successfully!**\n\n` +
          `🆔 **Task ID:** ${parameters.taskId}\n` +
          (parameters.subject ? `📋 **Subject:** ${parameters.subject}\n` : "") +
          (parameters.status ? `📊 **Status:** ${parameters.status}\n` : "") +
          (parameters.priority ? `⭐ **Priority:** ${parameters.priority}\n` : "") +
          (parameters.activityDate ? `📅 **Activity Date:** ${parameters.activityDate}\n` : "") +
          (parameters.description ? `📝 **Description:** ${parameters.description}\n` : "")
        )
      );
      return `Successfully updated Salesforce task ${parameters.taskId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to update task: ${response.message || "Unknown error"}.`)
      );
      return `Failed to update task: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error updating Salesforce task:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error updating task: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("DeleteSalesforceTask", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    if (!parameters.taskId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Task ID to delete a task.")
      );
      return "Missing required parameters";
    }

    console.log(`Deleting task ${parameters.taskId} from Salesforce CRM...`);

    const response = await deleteSalesforceTask(context, state, parameters.taskId);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(`✅ **Task Deleted Successfully!**\n\n🆔 **Task ID:** ${parameters.taskId}`)
      );
      return `Successfully deleted task ${parameters.taskId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to delete task: ${response.message || "Unknown error"}.`)
      );
      return `Failed to delete task: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error deleting Salesforce task:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error deleting task: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

// ======================= ACCOUNTS =======================

app.ai.action("GetSalesforceAccounts", async (context, state, parameters) => {
  console.log("GetSalesforceAccounts action called with parameters:", parameters);
  try {
    initializeConversationState(state);
    
    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, Name, Type, Industry, Phone, Website FROM Account ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const axios = require("axios");
    
    const headers = {
      Authorization: `Bearer ${config.salesforceAccessToken}`,
      "Content-Type": "application/json",
    };

    const response = await axios.get(
      `https://orgfarm-5a7d798f5f-dev-ed.develop.lightning.force.com/services/data/v60.0/query?q=${encodeURIComponent(query)}`,
      { headers }
    );

    const records = response.data.records || [];

    if (records.length === 0) {
      await context.sendActivity(MessageFactory.text("📊 No accounts found in your Salesforce CRM."));
      return "No accounts found";
    }

    const formattedAccounts = records.map((a) => ({
      id: a.Id,
      name: a.Name || "—",
      type: a.Type || "—",
      industry: a.Industry || "—",
      phone: a.Phone || "—",
      website: a.Website || "—",
    }));

    state.conversation.formattedAccounts = JSON.stringify(formattedAccounts, null, 2);
    return `Retrieved ${records.length} accounts successfully`;
  } catch (error) {
    console.error("Error fetching Salesforce accounts:", error);
    const errorMessage = error.response?.data?.[0]?.message || error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error retrieving accounts: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("CreateSalesforceAccount", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const name = parameters.name;

    if (!name) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `name`.")
      );
      return "Missing required parameters";
    }

    const response = await createSalesforceAccount(context, state, { name });

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Account Created Successfully!**\n\n` +
          `🆔 **Account ID:** ${response.id}\n` +
          `🏢 **Name:** ${name}`
        )
      );
      return `Successfully created Salesforce account ${name}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to create account: ${response.message || "Unknown error"}.`)
      );
      return `Failed to create account: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error creating Salesforce account:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error creating account: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("UpdateSalesforceAccount", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    if (!parameters.accountId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Account ID to update an account.")
      );
      return "Missing required parameters";
    }

    const updateFields = {
      ...(parameters.name && { Name: parameters.name }),
      ...(parameters.type && { Type: parameters.type }),
      ...(parameters.industry && { Industry: parameters.industry }),
      ...(parameters.phone && { Phone: parameters.phone }),
      ...(parameters.website && { Website: parameters.website }),
      ...(parameters.description && { Description: parameters.description }),
    };

    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity(
        MessageFactory.text("❌ No fields provided to update. Please include at least one field like name, type, etc.")
      );
      return "No fields provided to update";
    }

    console.log(`Updating account ${parameters.accountId} in Salesforce CRM...`);

    const response = await updateSalesforceAccount(context, state, parameters.accountId, updateFields);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Account Updated Successfully!**\n\n` +
          `🆔 **Account ID:** ${parameters.accountId}\n` +
          (parameters.name ? `🏢 **Name:** ${parameters.name}\n` : "") +
          (parameters.type ? `📋 **Type:** ${parameters.type}\n` : "") +
          (parameters.industry ? `🏭 **Industry:** ${parameters.industry}\n` : "") +
          (parameters.phone ? `📱 **Phone:** ${parameters.phone}\n` : "") +
          (parameters.website ? `🌐 **Website:** ${parameters.website}\n` : "") +
          (parameters.description ? `📝 **Description:** ${parameters.description}\n` : "")
        )
      );
      return `Successfully updated Salesforce account ${parameters.accountId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to update account: ${response.message || "Unknown error"}.`)
      );
      return `Failed to update account: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error updating Salesforce account:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error updating account: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("DeleteSalesforceAccount", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    if (!parameters.accountId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Account ID to delete an account.")
      );
      return "Missing required parameters";
    }

    console.log(`Deleting account ${parameters.accountId} from Salesforce CRM...`);

    const response = await deleteSalesforceAccount(context, state, parameters.accountId);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(`✅ **Account Deleted Successfully!**\n\n🆔 **Account ID:** ${parameters.accountId}`)
      );
      return `Successfully deleted account ${parameters.accountId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to delete account: ${response.message || "Unknown error"}.`)
      );
      return `Failed to delete account: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error deleting Salesforce account:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error deleting account: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

// ======================= CONTACTS =======================

app.ai.action("GetSalesforceContacts", async (context, state, parameters) => {
  console.log("GetSalesforceContacts action called with parameters:", parameters);
  try {
    initializeConversationState(state);
    
    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, FirstName, LastName, Email, Phone, Title, AccountId, Account.Name FROM Contact ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const axios = require("axios");
    
    const headers = {
      Authorization: `Bearer ${config.salesforceAccessToken}`,
      "Content-Type": "application/json",
    };

    const response = await axios.get(
      `https://orgfarm-5a7d798f5f-dev-ed.develop.lightning.force.com/services/data/v60.0/query?q=${encodeURIComponent(query)}`,
      { headers }
    );

    const records = response.data.records || [];

    if (records.length === 0) {
      await context.sendActivity(MessageFactory.text("📊 No contacts found in your Salesforce CRM."));
      return "No contacts found";
    }

    const formattedContacts = records.map((c) => ({
      id: c.Id,
      name: `${c.FirstName || ''} ${c.LastName || ''}`.trim() || "—",
      email: c.Email || "—",
      phone: c.Phone || "—",
      title: c.Title || "—",
      accountName: c.Account?.Name || "—",
    }));

    state.conversation.formattedContacts = JSON.stringify(formattedContacts, null, 2);
    return `Retrieved ${records.length} contacts successfully`;
  } catch (error) {
    console.error("Error fetching Salesforce contacts:", error);
    const errorMessage = error.response?.data?.[0]?.message || error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error retrieving contacts: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("CreateSalesforceContact", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const firstName = parameters.firstName;
    const lastName = parameters.lastName;

    if (!firstName || !lastName) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `firstName` and `lastName`.")
      );
      return "Missing required parameters";
    }

    const response = await createSalesforceContact(context, state, { firstName, lastName });

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Contact Created Successfully!**\n\n` +
          `🆔 **Contact ID:** ${response.id}\n` +
          `👤 **Name:** ${firstName} ${lastName}`
        )
      );
      return `Successfully created Salesforce contact ${firstName} ${lastName}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to create contact: ${response.message || "Unknown error"}.`)
      );
      return `Failed to create contact: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error creating Salesforce contact:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error creating contact: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("UpdateSalesforceContact", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    if (!parameters.contactId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Contact ID to update a contact.")
      );
      return "Missing required parameters";
    }

    const updateFields = {
      ...(parameters.firstName && { FirstName: parameters.firstName }),
      ...(parameters.lastName && { LastName: parameters.lastName }),
      ...(parameters.email && { Email: parameters.email }),
      ...(parameters.phone && { Phone: parameters.phone }),
      ...(parameters.title && { Title: parameters.title }),
      ...(parameters.accountId && { AccountId: parameters.accountId }),
      ...(parameters.department && { Department: parameters.department }),
    };

    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity(
        MessageFactory.text("❌ No fields provided to update. Please include at least one field like firstName, email, etc.")
      );
      return "No fields provided to update";
    }

    console.log(`Updating contact ${parameters.contactId} in Salesforce CRM...`);

    const response = await updateSalesforceContact(context, state, parameters.contactId, updateFields);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Contact Updated Successfully!**\n\n` +
          `🆔 **Contact ID:** ${parameters.contactId}\n` +
          (parameters.firstName ? `👤 **First Name:** ${parameters.firstName}\n` : "") +
          (parameters.lastName ? `👤 **Last Name:** ${parameters.lastName}\n` : "") +
          (parameters.email ? `📧 **Email:** ${parameters.email}\n` : "") +
          (parameters.phone ? `📱 **Phone:** ${parameters.phone}\n` : "") +
          (parameters.title ? `💼 **Title:** ${parameters.title}\n` : "") +
          (parameters.department ? `🏢 **Department:** ${parameters.department}\n` : "")
        )
      );
      return `Successfully updated Salesforce contact ${parameters.contactId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to update contact: ${response.message || "Unknown error"}.`)
      );
      return `Failed to update contact: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error updating Salesforce contact:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error updating contact: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});

app.ai.action("DeleteSalesforceContact", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    if (!parameters.contactId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Contact ID to delete a contact.")
      );
      return "Missing required parameters";
    }

    console.log(`Deleting contact ${parameters.contactId} from Salesforce CRM...`);

    const response = await deleteSalesforceContact(context, state, parameters.contactId);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(`✅ **Contact Deleted Successfully!**\n\n🆔 **Contact ID:** ${parameters.contactId}`)
      );
      return `Successfully deleted contact ${parameters.contactId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to delete contact: ${response.message || "Unknown error"}.`)
      );
      return `Failed to delete contact: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error deleting Salesforce contact:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error deleting contact: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});


app.activity(ActivityTypes.Message, async (context, state) => {
  try {
    // initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    console.log("context.activity:", context.activity)
    if (context.activity.text?.startsWith("/")) {
      return;
    }

    const actionType = context?.activity?.value?.action ?? false

    // if (await valiAction(actionType)) {
    //   await checkActionType(context, state, actionType)
    //   return
    // }

    // state.conversation.userId = userId;
    // try {
    //   const token = await getUserToken(teamsChatId);
    //   state.conversation.isAuthenticated = !!(token && token.accessToken);
    // } catch {
    //   state.conversation.isAuthenticated = false;
    // }

    console.log(`Processing message from user: ${userId}, authenticated: ${state.conversation.isAuthenticated}, state:`, state.conversation);
    await context.sendActivity(
        MessageFactory.text(
          "👋 **Welcome to your AI-powered Salesforce CRM assistant!**\n\n" +
          "To get started, we need to connect your SalesForce CRM account. " +
          "Once connected, you can ask me questions about your CRM data using natural language!"
        )
      );
    // if (!state.conversation.isAuthenticated) {
    //   await context.sendActivity(
    //     MessageFactory.text(
    //       "👋 **Welcome to your AI-powered Zoho CRM assistant!**\n\n" +
    //       "To get started, we need to connect your Zoho CRM account. " +
    //       "Once connected, you can ask me questions about your CRM data using natural language!"
    //     )
    //   );
    //   await sendZohoLoginCard(context, userId);
    //   return;
    // }

    // state.temp = state.temp || {};
    // state.temp.input = context.activity.text;

    console.log("Running AI with input:", state.temp.input);
    await app.ai.run(context, state);
    console.log("state updated:", state.conversation);
    if (state.temp?.plan?.commands) {
      for (const cmd of state.temp.plan.commands) {
        if (cmd.type === "SAY") {
          await context.sendActivity(MessageFactory.text(cmd.response));
          console.log("SAY response sent to Teams:", cmd.response);
        }
      }
    }
  } catch (error) {
    console.error("Error in message handler:", error);
    throw error;
  }
});


module.exports = app;
