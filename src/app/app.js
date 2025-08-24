const { MemoryStorage, MessageFactory, CardFactory, ActivityTypes } = require("botbuilder");
const path = require("path");
const config = require("../config");
const chrono = require("chrono-node");
const moment = require("moment-timezone");
const { getUserToken } = require("../cosmos");
const { httpRequest } = require("../httprequest");
const { deleteUserToken } = require("../cosmos");
const { getOutlookLoginCard } = require("../adaptiveCards/outlookLoginCard");
const { getOutlookToken, initializeConversationStateOutlook, getRecentEmails } = require('../oauthhandler/salesforceauth');

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
  createSalesforceMeeting,
  findSalesforceContactOrAccount,
  findSalesforceMeeting,
  updateSalesforceMeeting,
  cancelSalesforceMeeting,
  generateOpportunityFromEmail
} = require("../salesforce");

const {
  getSalesforceLoginCard
} = require("../adaptiveCards/salesforceLoginCard");


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

    const response = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(query)}`, "GET");

    const records = response.data.records || [];

    if (records.length === 0) {
      // await context.sendActivity(
      //   MessageFactory.text("📊 No leads found in your Salesforce CRM.")
      // );
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
    // await context.sendActivity(
    //   MessageFactory.text(`❌ Error retrieving leads: ${errorMessage}. Please try again.`)
    // );
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
    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Lead`, "POST", { firstName, lastName, company });
    // const response = await createSalesforceLead(context, state, { firstName, lastName, company });

    if (response.data.id) {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Lead Created Successfully!**\n\n` +
          `🆔 **Lead ID:** ${response.data.id}\n` +
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
      // const emailResponse = await axios.get(
      //   `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(emailQuery)}`,
      //   { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      // );

      const emailResponse = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(emailQuery)}`, "GET");

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

      // const nameResponse = await axios.get(
      //   `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
      //   { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      // );
      const nameResponse = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`, "GET");

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
    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Lead/${leadId}`, "PATCH", updateFields);
    await context.sendActivity(`✅ Lead updated successfully (ID: ${leadId}).`);
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

      // const headers = {
      //   Authorization: `Bearer ${config.salesforceAccessToken}`,
      //   "Content-Type": "application/json",
      // };

      // const queryUrl = `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`;
      // const lookupResponse = await axios.get(queryUrl, { headers });
      const lookupResponse = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`, "GET");

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
    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Lead/${leadId}`, "DELETE");
    await context.sendActivity(
      MessageFactory.text(`✅ **Lead Deleted Successfully!**\n\n🆔 **Lead ID:** ${leadId}`)
    );
    return `Successfully deleted lead ${leadId}`;

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
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, Name, StageName, Amount, CloseDate, AccountId, Account.Name FROM Opportunity ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const response = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(query)}`, "GET");


    const records = response.data.records || [];

    if (records.length === 0) {
      // await context.sendActivity(MessageFactory.text("📊 No opportunities found in your Salesforce CRM."));
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
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const name = parameters.name;
    const stageName = parameters.stageName || "Prospecting";
    const closeDate = parameters.closeDate;

    if (!name || !closeDate) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `name` and `closeDate`.")
      );
      return "Missing required parameters";
    }

    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Opportunity`, "POST", { Name: name, StageName: stageName, CloseDate: closeDate });


    if (response.data.id) {
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
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const config = require("../config");
    const axios = require("axios");

    let opportunityId = parameters.opportunityId;

    // 1. Try finding opportunity by name first if no ID provided
    if (!opportunityId && (parameters.opportunityName || parameters.name)) {
      const opportunityName = parameters.opportunityName || parameters.name;
      const nameQuery = `SELECT Id, Name FROM Opportunity WHERE Name LIKE '%${opportunityName}%' LIMIT 200`;

      // const nameResponse = await axios.get(
      //   `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
      //   { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      // );
      const nameResponse = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`, "GET");

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
    const response = await httpRequest(teamsChatId, `/sobjectsOpportunity/${opportunityId}`, "PATCH", updateFields);
    await context.sendActivity(`✅ Opportunity updated successfully (ID: ${opportunityId}).`);
  } catch (error) {
    console.error("Error in UpdateSalesforceOpportunity:", error);
    await context.sendActivity(`❌ Error: ${error.message || "Unknown error"}`);
  }
});


app.ai.action("DeleteSalesforceOpportunity", async (context, state, parameters) => {
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    // 1. If no opportunityId is given, try to find it by name
    let opportunityId = parameters.opportunityId;
    if (parameters.opportunityName || parameters.name) {
      const nameToSearch = parameters.opportunityName || parameters.name;

      const nameQuery = `SELECT Id, Name FROM Opportunity WHERE Name LIKE '%${nameToSearch}%' LIMIT 200`;

      const lookupResponse = await httpRequest(
        teamsChatId,
        `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
        "GET"
      );

      if (lookupResponse.data.records.length === 0) {
        await context.sendActivity(
          MessageFactory.text(`❌ No Salesforce opportunity found matching the provided name.`)
        );
        return "No opportunity found by name";
      }
      if (lookupResponse.data.records.length > 1) {
        const opportunityList = lookupResponse.data.records
          .map((opp, index) => `${index + 1}. ${opp.Name} (${opp.Id})`)
          .join("\n");

        await context.sendActivity(
          MessageFactory.text(
            `⚠ Multiple opportunities found matching the provided name:\n\n${opportunityList}\n\n` +
            `Please refine your search.`
          )
        );
        return "Multiple opportunities found by name";
      }

      opportunityId = lookupResponse.data.records[0].Id;
    }

    if (!opportunityId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need either: Opportunity ID or a name to find the opportunity.")
      );
      return "Missing required parameters";
    }

    console.log(`Deleting opportunity ${opportunityId} from Salesforce CRM...`);
    const response = await httpRequest(
      teamsChatId,
      `/services/data/v59.0/sobjects/Opportunity/${opportunityId}`,
      "DELETE"
    );

    if (response.status !== 204 && response.success === false) {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to delete opportunity: ${response.message || "Unknown error"}.`)
      );
      return `Failed to delete opportunity: ${response.message || "Unknown error"}`;
    }

    await context.sendActivity(
      MessageFactory.text(`✅ **Opportunity Deleted Successfully!**\n\n🆔 **Opportunity ID:** ${opportunityId}`)
    );
    return `Successfully deleted opportunity ${opportunityId}`;

  } catch (error) {
    console.error("Error deleting Salesforce opportunity:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error deleting opportunity: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});





// ======================= ACCOUNTS =======================

app.ai.action("GetSalesforceAccounts", async (context, state, parameters) => {
  console.log("GetSalesforceAccounts action called with parameters:", parameters);
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, Name, Type, Industry, Phone, Website FROM Account ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const response = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(query)}`, "GET");

    const records = response.data.records || [];

    if (records.length === 0) {
      // await context.sendActivity(MessageFactory.text("📊 No accounts found in your Salesforce CRM."));
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

    // Create Adaptive Card
    const adaptiveCard = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "🏢 Your Salesforce Accounts",
          size: "Large",
          weight: "Bolder",
          color: "Light"
        },
        {
          type: "TextBlock",
          text: `Found ${records.length} accounts`,
          size: "Medium",
          color: "Good",
          spacing: "Small"
        },
        ...formattedAccounts.slice(0, 10).map((account, index) => ({
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
                      text: `${index + 1}. ${account.name}`,
                      weight: "Bolder",
                      size: "Medium",
                      wrap: true,
                      color: "Light"
                    },
                    {
                      type: "TextBlock",
                      text: `Industry: ${account.industry}`,
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
                      text: account.type,
                      weight: "Bolder",
                      size: "Small",
                      color: "Accent",
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
                      text: `📞 ${account.phone}`,
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
                      text: account.website !== "—" ? `🌐 ${account.website}` : "🌐 —",
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

    // If there are more than 10 accounts, add a note
    if (formattedAccounts.length > 10) {
      adaptiveCard.body.push({
        type: "TextBlock",
        text: `... and ${formattedAccounts.length - 10} more accounts`,
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
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const name = parameters.name;

    if (!name) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `name`.")
      );
      return "Missing required parameters";
    }

    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Account`, "POST", { Name: name });


    if (response.data.id) {
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
    console.log("UpdateSalesforceAccount called with:", parameters);
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const config = require("../config");
    const axios = require("axios");

    let accountId = parameters.accountId;

    // 1. Try finding account by name first if no ID provided
    if (!accountId && (parameters.accountName || parameters.name)) {
      const accountName = parameters.accountName || parameters.name;
      const nameQuery = `SELECT Id, Name FROM Account WHERE Name LIKE '%${accountName}%' LIMIT 200`;

      // const nameResponse = await axios.get(
      //   `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
      //   { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      // );
      const nameResponse = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`, "GET");

      if (nameResponse.data.records.length === 1) {
        accountId = nameResponse.data.records[0].Id;
      } else if (nameResponse.data.records.length > 1) {
        await context.sendActivity("⚠ Multiple accounts found by name. Please refine search.");
        return;
      } else if (nameResponse.data.records.length === 0) {
        await context.sendActivity("❌ No matching Salesforce account found.");
        return;
      }
    }

    if (!accountId) {
      await context.sendActivity("❌ Could not find a matching account by provided details.");
      return;
    }

    // 2. Build update object
    const updateFields = {
      ...(parameters.newName && { Name: parameters.newName }),
      ...(parameters.type && { Type: parameters.type }),
      ...(parameters.industry && { Industry: parameters.industry }),
      ...(parameters.phone && { Phone: parameters.phone }),
      ...(parameters.website && { Website: parameters.website }),
      ...(parameters.description && { Description: parameters.description })
    };

    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity("❌ No fields provided to update.");
      return;
    }

    // console.log(`Updating account ${accountId} with fields:`, updateFields);
    // const { updateSalesforceAccount } = require("../salesforce");
    // const response = await updateSalesforceAccount(context, state, accountId, updateFields);
    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Account/${accountId}`, "PATCH", updateFields);
    await context.sendActivity(`✅ Account updated successfully (ID: ${accountId}).`);
  } catch (error) {
    console.error("Error in UpdateSalesforceAccount:", error);
    await context.sendActivity(`❌ Error: ${error.message || "Unknown error"}`);
  }
});


app.ai.action("DeleteSalesforceAccount", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    // 1. If no accountId is given, try to find it by name
    let accountId = parameters.accountId;
    if (parameters.accountName || parameters.name) {
      const nameValue = parameters.accountName || parameters.name;
      const nameQuery = `SELECT Id, Name FROM Account WHERE Name LIKE '%${nameValue}%' LIMIT 200`;

      const lookupResponse = await httpRequest(
        teamsChatId,
        `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
        "GET"
      );

      if (lookupResponse.data.records.length === 0) {
        await context.sendActivity(
          MessageFactory.text(`❌ No Salesforce account found matching the provided name.`)
        );
        return "No account found by name";
      }

      if (lookupResponse.data.records.length > 1) {
        const accountList = lookupResponse.data.records
          .map((acc, index) => `${index + 1}. ${acc.Name} (${acc.Id})`)
          .join("\n");

        await context.sendActivity(
          MessageFactory.text(
            `⚠ Multiple accounts found matching the provided name:\n\n${accountList}\n\nPlease refine your search.`
          )
        );
        return "Multiple accounts found by name";
      }

      accountId = lookupResponse.data.records[0].Id;
    }

    // 2. If still no accountId, return error
    if (!accountId) {
      await context.sendActivity(
        MessageFactory.text(
          "❌ Missing required information. I need either: Account ID or an Account Name to find the account."
        )
      );
      return "Missing required parameters";
    }

    // 3. Delete account
    console.log(`Deleting account ${accountId} from Salesforce CRM...`);
    const response = await httpRequest(
      teamsChatId,
      `/services/data/v59.0/sobjects/Account/${accountId}`,
      "DELETE"
    );

    await context.sendActivity(
      MessageFactory.text(
        `✅ **Account Deleted Successfully!**\n\n🆔 **Account ID:** ${accountId}`
      )
    );
    return `Successfully deleted account ${accountId}`;
  } catch (error) {
    console.error("Error deleting Salesforce account:", error);
    const errorMessage = error.message || "Unknown error";

    await context.sendActivity(
      MessageFactory.text(
        `❌ Error deleting account: ${errorMessage}. Please try again.`
      )
    );
    return `Error occurred: ${errorMessage}`;
  }
});



// ======================= TASKS =======================

app.ai.action("GetSalesforceTasks", async (context, state, parameters) => {
  console.log("GetSalesforceTasks action called with parameters:", parameters);

  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, Subject, Status, Priority, ActivityDate, Description 
                   FROM Task 
                   ORDER BY CreatedDate DESC 
                   LIMIT ${limit}`;

    const response = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(query)}`, "GET");

    const records = response.data.records || [];

    if (records.length === 0) {
      await context.sendActivity({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: {
              type: "AdaptiveCard",
              version: "1.5",
              body: [
                {
                  type: "TextBlock",
                  text: "📊 No tasks found",
                  weight: "Bolder",
                  size: "Medium",
                  color: "Attention"
                }
              ]
            }
          }
        ]
      });
      return "No tasks found";
    }

    const taskItems = records.map((t) => {
      return {
        type: "Container",
        style: "emphasis",
        items: [
          { type: "TextBlock", text: `**${t.Subject || "—"}**`, weight: "Bolder", size: "Medium" },
          { type: "TextBlock", text: `🗓 Due: ${t.ActivityDate || "—"}`, spacing: "Small" },
          { type: "TextBlock", text: `📌 Status: ${t.Status || "—"}`, spacing: "Small" },
          { type: "TextBlock", text: `⚡ Priority: ${t.Priority || "—"}`, spacing: "Small" },
          { type: "TextBlock", text: t.Description || "No description", wrap: true, spacing: "Small" }
        ]
      };
    });

    const adaptiveCard = {
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        {
          type: "TextBlock",
          text: `📋 Salesforce Tasks (${records.length})`,
          weight: "Bolder",
          size: "Large",
          separator: true
        },
        {
          type: "Container",
          items: taskItems
        }
      ],
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
    };

    await context.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: adaptiveCard
        }
      ]
    });

    state.conversation.formattedTasks = JSON.stringify(records, null, 2);
    return `Retrieved ${records.length} tasks successfully`;
  } catch (error) {
    console.error("Error fetching Salesforce tasks:", error);
    const errorMessage = error.response?.data?.[0]?.message || error.message || "Unknown error";

    const errorCard = {
      type: "AdaptiveCard",
      version: "1.5",
      body: [
        {
          type: "TextBlock",
          text: `❌ Error retrieving tasks: ${errorMessage}`,
          weight: "Bolder",
          color: "Attention",
          wrap: true
        }
      ]
    };

    await context.sendActivity({
      type: "message",
      attachments: [
        {
          contentType: "application/vnd.microsoft.card.adaptive",
          content: errorCard
        }
      ]
    });

    return `Error occurred: ${errorMessage}`;
  }
});


app.ai.action("CreateSalesforceTask", async (context, state, parameters) => {
  try {
    console.log("CreateSalesforceTask action called with parameters:", parameters);
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const subject = parameters.subject;
    const status = parameters.status || "Not Started";

    if (!subject) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `subject`.")
      );
      return "Missing required parameters";
    }

    const response = await httpCreateRequest(teamsChatId, `/services/data/v59.0/sobjects/Task`, "POST", { Subject: subject, Status: status });


    if (response && response.id) {
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
    console.log("UpdateSalesforceTask called with:", parameters);
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const config = require("../config");
    const axios = require("axios");

    let taskId = parameters.taskId;

    // 1. Try finding task by subject if no ID provided
    if (!taskId && parameters.taskSubject) {
      const taskSubject = parameters.taskSubject;
      const nameQuery = `SELECT Id, Subject FROM Task WHERE Subject LIKE '%${taskSubject}%' LIMIT 200`;

      // const nameResponse = await axios.get(
      //   `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
      //   { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      // );
      const nameResponse = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`, "GET");

      console.log("Name query response:", nameResponse.data.records);
      if (nameResponse.data.records.length === 1) {
        taskId = nameResponse.data.records[0].Id;
      } else if (nameResponse.data.records.length > 1) {
        await context.sendActivity("⚠ Multiple tasks found with that subject. Please refine your search.");
        return;
      } else if (nameResponse.data.records.length === 0) {
        await context.sendActivity("❌ No matching Salesforce task found.");
        return;
      }
    }

    if (!taskId) {
      await context.sendActivity("❌ Could not find a matching task by provided details.");
      return;
    }

    // 2. Build update object
    const updateFields = {
      ...(parameters.subject && { Subject: parameters.subject }),
      ...(parameters.status && { Status: parameters.status }),
      ...(parameters.priority && { Priority: parameters.priority }),
      ...(parameters.activityDate && { ActivityDate: parameters.activityDate }),
      ...(parameters.description && { Description: parameters.description })
    };

    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity("❌ No fields provided to update.");
      return;
    }

    // console.log(`Updating task ${taskId} with fields:`, updateFields);
    // const { updateSalesforceTask } = require("../salesforce");
    // const response = await updateSalesforceTask(context, state, taskId, updateFields);
    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Task/${taskId}`, "PATCH", updateFields);
    await context.sendActivity(`✅ Task updated successfully (ID: ${taskId}).`);
  } catch (error) {
    console.error("Error in UpdateSalesforceTask:", error.response.data);
    await context.sendActivity(`❌ Error: ${error.message || "Unknown error"}`);
  }
});


app.ai.action("DeleteSalesforceTask", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    // 1. If no taskId is given, try to find it by subject
    let taskId = parameters.taskId;
    if (parameters.taskSubject || parameters.subject) {
      const subjectValue = parameters.taskSubject || parameters.subject;
      const subjectQuery = `SELECT Id, Subject FROM Task WHERE Subject LIKE '%${subjectValue}%' LIMIT 200`;

      const lookupResponse = await httpRequest(
        teamsChatId,
        `/services/data/v59.0/query?q=${encodeURIComponent(subjectQuery)}`,
        "GET"
      );

      if (lookupResponse.data.records.length === 0) {
        await context.sendActivity(
          MessageFactory.text(`❌ No Salesforce task found matching the provided subject.`)
        );
        return "No task found by subject";
      }

      if (lookupResponse.data.records.length > 1) {
        const taskList = lookupResponse.data.records
          .map((t, index) => `${index + 1}. ${t.Subject} (${t.Id})`)
          .join("\n");

        await context.sendActivity(
          MessageFactory.text(
            `⚠ Multiple tasks found matching the provided subject:\n\n${taskList}\n\nPlease refine your search.`
          )
        );
        return "Multiple tasks found by subject";
      }

      taskId = lookupResponse.data.records[0].Id;
    }

    // 2. If still no taskId, return error
    if (!taskId) {
      await context.sendActivity(
        MessageFactory.text(
          "❌ Missing required information. I need either: Task ID or a Task Subject to find the task."
        )
      );
      return "Missing required parameters";
    }

    // 3. Delete task
    console.log(`Deleting task ${taskId} from Salesforce CRM...`);
    const response = await httpRequest(
      teamsChatId,
      `/services/data/v59.0/sobjects/Task/${taskId}`,
      "DELETE"
    );

    await context.sendActivity(
      MessageFactory.text(
        `✅ **Task Deleted Successfully!**\n\n🆔 **Task ID:** ${taskId}`
      )
    );
    return `Successfully deleted task ${taskId}`;
  } catch (error) {
    console.error("Error deleting Salesforce task:", error);
    const errorMessage = error.message || "Unknown error";

    await context.sendActivity(
      MessageFactory.text(
        `❌ Error deleting task: ${errorMessage}. Please try again.`
      )
    );
    return `Error occurred: ${errorMessage}`;
  }
});






// ======================= CONTACTS =======================

app.ai.action("GetSalesforceContacts", async (context, state, parameters) => {
  console.log("GetSalesforceContacts action called with parameters:", parameters);
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, FirstName, LastName, Email, Phone, Title, AccountId, Account.Name FROM Contact ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const response = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(query)}`, "GET");

    const records = response.data.records || [];

    if (records.length === 0) {
      // await context.sendActivity(MessageFactory.text("📊 No contacts found in your Salesforce CRM."));
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

    // Create Adaptive Card
    const adaptiveCard = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "📇 Your Salesforce Contacts",
          size: "Large",
          weight: "Bolder",
          color: "Light"
        },
        {
          type: "TextBlock",
          text: `Found ${records.length} contacts`,
          size: "Medium",
          color: "Good",
          spacing: "Small"
        },
        ...formattedContacts.slice(0, 10).map((contact, index) => ({
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
                      text: `${index + 1}. ${contact.name}`,
                      weight: "Bolder",
                      size: "Medium",
                      wrap: true,
                      color: "Light"
                    },
                    {
                      type: "TextBlock",
                      text: `Title: ${contact.title}`,
                      color: "Light",
                      size: "Small",
                      spacing: "None"
                    },
                    {
                      type: "TextBlock",
                      text: `Account: ${contact.accountName}`,
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
                      text: contact.email,
                      size: "Small",
                      color: "Accent",
                      wrap: true,
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
                      text: `📞 ${contact.phone}`,
                      size: "Small",
                      color: "Light"
                    }
                  ]
                }
              ]
            }
          ]
        }))
      ]
    };

    // If there are more than 10 contacts, add a note
    if (formattedContacts.length > 10) {
      adaptiveCard.body.push({
        type: "TextBlock",
        text: `... and ${formattedContacts.length - 10} more contacts`,
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
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const firstName = parameters.firstName;
    const lastName = parameters.lastName;

    if (!firstName || !lastName) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required fields. Please provide `firstName` and `lastName`.")
      );
      return "Missing required parameters";
    }

    const response = await httpCreateRequest(teamsChatId, `/services/data/v57.0/sobjects/Contact`, "POST", { FirstName: firstName, LastName: lastName });

    if (response && response.data) {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Contact Created Successfully!**\n\n` +
          `🆔 **Contact ID:** ${response.data.id}\n` +
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
    console.log("UpdateSalesforceContact called with:", parameters);
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const config = require("../config");
    const axios = require("axios");

    let contactId = parameters.contactId;

    // 1. Try finding contact by name if no ID provided
    if (!contactId && (parameters.contactName || parameters.name || parameters.firstName || parameters.lastName)) {
      const contactName = parameters.contactName || parameters.name || "";
      const firstName = parameters.firstName || "";
      const lastName = parameters.lastName || "";

      let nameQuery = "";
      if (contactName) {
        // Search by full name string
        nameQuery = `SELECT Id, FirstName, LastName FROM Contact WHERE Name LIKE '%${contactName}%' LIMIT 200`;
      } else if (firstName && lastName) {
        nameQuery = `SELECT Id, FirstName, LastName FROM Contact WHERE FirstName LIKE '%${firstName}%' AND LastName LIKE '%${lastName}%' LIMIT 200`;
      } else if (firstName) {
        nameQuery = `SELECT Id, FirstName, LastName FROM Contact WHERE FirstName LIKE '%${firstName}%' LIMIT 200`;
      } else if (lastName) {
        nameQuery = `SELECT Id, FirstName, LastName FROM Contact WHERE LastName LIKE '%${lastName}%' LIMIT 200`;
      }

      // const nameResponse = await axios.get(
      //   `https://orgfarm-5a7d798f5f-dev-ed.develop.my.salesforce.com/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
      //   { headers: { Authorization: `Bearer ${config.salesforceAccessToken}` } }
      // );
      const nameResponse = await httpRequest(teamsChatId, `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`, "GET");

      if (nameResponse.data.records.length === 1) {
        contactId = nameResponse.data.records[0].Id;
      } else if (nameResponse.data.records.length > 1) {
        await context.sendActivity("⚠ Multiple contacts found by that name. Please refine search.");
        return;
      } else if (nameResponse.data.records.length === 0) {
        await context.sendActivity("❌ No matching Salesforce contact found.");
        return;
      }
    }

    if (!contactId) {
      await context.sendActivity("❌ Could not find a matching contact with the provided details.");
      return;
    }

    // 2. Build update object
    const updateFields = {
      ...(parameters.firstName && { FirstName: parameters.firstName }),
      ...(parameters.lastName && { LastName: parameters.lastName }),
      ...(parameters.email && { Email: parameters.email }),
      ...(parameters.phone && { Phone: parameters.phone }),
      ...(parameters.title && { Title: parameters.title }),
      ...(parameters.accountId && { AccountId: parameters.accountId }),
      ...(parameters.department && { Department: parameters.department })
    };

    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity("❌ No fields provided to update. Please include at least one field like firstName, email, etc.");
      return;
    }

    // console.log(`Updating contact ${contactId} with fields:`, updateFields);
    // const { updateSalesforceContact } = require("../salesforce");
    // const response = await updateSalesforceContact(context, state, contactId, updateFields);
    const response = await httpRequest(teamsChatId, `/services/data/v59.0/sobjects/Contact/${contactId}`, "PATCH", updateFields);
    await context.sendActivity(`✅ Contact updated successfully (ID: ${contactId}).`);
  } catch (error) {
    console.error("Error in UpdateSalesforceContact:", error);
    await context.sendActivity(`❌ Error: ${error.message || "Unknown error"}`);
  }
});


app.ai.action("DeleteSalesforceContact", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    // 1. If no contactId is given, try to find it by name
    let contactId = parameters.contactId;
    if (parameters.contactName || parameters.name) {
      const nameValue = parameters.contactName || parameters.name;
      const nameQuery = `SELECT Id, FirstName, LastName FROM Contact WHERE Name LIKE '%${nameValue}%' LIMIT 200`;

      const lookupResponse = await httpRequest(
        teamsChatId,
        `/services/data/v59.0/query?q=${encodeURIComponent(nameQuery)}`,
        "GET"
      );

      if (lookupResponse.data.records.length === 0) {
        await context.sendActivity(
          MessageFactory.text(`❌ No Salesforce contact found matching the provided name.`)
        );
        return "No contact found by name";
      }

      if (lookupResponse.data.records.length > 1) {
        const contactList = lookupResponse.data.records
          .map(
            (c, index) =>
              `${index + 1}. ${c.FirstName || ""} ${c.LastName || ""} (${c.Id})`
          )
          .join("\n");

        await context.sendActivity(
          MessageFactory.text(
            `⚠ Multiple contacts found matching the provided name:\n\n${contactList}\n\nPlease refine your search.`
          )
        );
        return "Multiple contacts found by name";
      }

      contactId = lookupResponse.data.records[0].Id;
    }

    // 2. If still no contactId, return error
    if (!contactId) {
      await context.sendActivity(
        MessageFactory.text(
          "❌ Missing required information. I need either: Contact ID or a Contact Name to find the contact."
        )
      );
      return "Missing required parameters";
    }

    // 3. Delete contact
    console.log(`Deleting contact ${contactId} from Salesforce CRM...`);
    await httpRequest(
      teamsChatId,
      `/services/data/v59.0/sobjects/Contact/${contactId}`,
      "DELETE"
    );

    await context.sendActivity(
      MessageFactory.text(
        `✅ **Contact Deleted Successfully!**\n\n🆔 **Contact ID:** ${contactId}`
      )
    );
    return `Successfully deleted contact ${contactId}`;
  } catch (error) {
    console.error("Error deleting Salesforce contact:", error);
    const errorMessage = error.message || "Unknown error";

    await context.sendActivity(
      MessageFactory.text(
        `❌ Error deleting contact: ${errorMessage}. Please try again.`
      )
    );
    return `Error occurred: ${errorMessage}`;
  }
});




app.ai.action("CreateSalesforceMeeting", async (context, state, parameters) => {
  try {
    initializeConversationState(state);

    const { subject, startDateTime, withName } = parameters;

    if (!subject || !startDateTime) {
      await context.sendActivity("❌ Please provide both the meeting subject and start date/time.");
      return;
    }

    // Get user's timezone
    const userTimeZone = "Asia/Kolkata"; // Hardcode to IST since you're in India

    // Parse the natural language datetime
    let parsedDate = chrono.parseDate(startDateTime, new Date(), { forwardDate: true });
    if (!parsedDate) {
      await context.sendActivity("❌ Could not understand the meeting start date/time.");
      return;
    }

    console.log("Original parsed date:", parsedDate);
    console.log("Original parsed date ISO:", parsedDate.toISOString());

    // THE KEY FIX: Create the datetime as if the user meant local time
    // Parse the time components from the original request
    const chronoResult = chrono.parse(startDateTime, new Date(), { forwardDate: true })[0];

    // Get the date and time components
    const year = chronoResult.start.get('year');
    const month = chronoResult.start.get('month') - 1; // JavaScript months are 0-based
    const day = chronoResult.start.get('day');
    const hour = chronoResult.start.get('hour') || 0;
    const minute = chronoResult.start.get('minute') || 0;

    console.log("Parsed components:", { year, month: month + 1, day, hour, minute });

    // Create moment in user's local timezone with these components
    const startMoment = moment.tz({ year, month, day, hour, minute }, userTimeZone);
    const endMoment = startMoment.clone().add(1, 'hour');

    console.log("Start moment in IST:", startMoment.format("YYYY-MM-DD HH:mm:ss Z"));
    console.log("End moment in IST:", endMoment.format("YYYY-MM-DD HH:mm:ss Z"));

    // Convert to UTC for Salesforce
    const startISO = startMoment.format('YYYY-MM-DDTHH:mm:ss.SSSZ');
    const endISO = endMoment.format('YYYY-MM-DDTHH:mm:ss.SSSZ');

    console.log("Start ISO for Salesforce (with IST offset):", startISO);
    console.log("End ISO for Salesforce (with IST offset):", endISO);

    let relatedRecordId = null;
    let relatedRecordType = null;

    // Contact/Account search logic
    if (withName) {
      const { findSalesforceContactOrAccount } = require("../salesforce");
      const found = await findSalesforceContactOrAccount(withName);

      if (found && found.Id) {
        relatedRecordId = found.Id;
        relatedRecordType = found.type;
      } else {
        if (parameters.withNameType === "Account") {
          const { createSalesforceAccount } = require("../salesforce");
          const accRes = await createSalesforceAccount({ Name: withName });
          relatedRecordId = accRes.Id;
          relatedRecordType = "Account";
        } else {
          const { createSalesforceContact } = require("../salesforce");
          const conRes = await createSalesforceContact({
            FirstName: withName.split(" ")[0],
            LastName: withName.split(" ")[1] || ""
          });
          relatedRecordId = conRes.Id;
          relatedRecordType = "Contact";
        }
      }
    }

    // Create the meeting
    const { createSalesforceMeeting } = require("../salesforce");
    const meetingRes = await createSalesforceMeeting({
      subject,
      startDateTime: startISO,
      endDateTime: endISO,
      whoId: relatedRecordType === "Contact" ? relatedRecordId : null,
      whatId: relatedRecordType === "Account" ? relatedRecordId : null
    });

    if (meetingRes.status === "success") {
      // Display in local time
      const localDisplay = startMoment.format("hh:mm A, DD MMM YYYY");
      await context.sendActivity(
        `✅ Meeting **"${subject}"** scheduled on ${localDisplay} with ${withName || "no specific person"}.`
      );
    } else {
      await context.sendActivity(`❌ Failed to schedule meeting: ${meetingRes.message}`);
    }

  } catch (err) {
    console.error("Error creating Salesforce meeting:", err);
    await context.sendActivity(`❌ Error creating meeting: ${err.message}`);
  }
});


// NEW ACTION: Update/Reschedule Meeting
app.ai.action("UpdateSalesforceMeeting", async (context, state, parameters) => {
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const { identifier, newDateTime, identifierType } = parameters;
    console.log("UpdateSalesforceMeeting called with:", teamsChatId);

    if (!identifier || !newDateTime) {
      await context.sendActivity("❌ Please provide both the meeting identifier (subject/time) and new date/time.");
      return;
    }

    // Get user's timezone
    const userTimeZone = "Asia/Kolkata";

    // Parse the new datetime
    let parsedNewDate = chrono.parseDate(newDateTime, new Date(), { forwardDate: true });
    if (!parsedNewDate) {
      await context.sendActivity("❌ Could not understand the new meeting date/time.");
      return;
    }

    // Parse new datetime components
    const chronoResult = chrono.parse(newDateTime, new Date(), { forwardDate: true })[0];
    const year = chronoResult.start.get('year');
    const month = chronoResult.start.get('month') - 1;
    const day = chronoResult.start.get('day');
    const hour = chronoResult.start.get('hour') || 0;
    const minute = chronoResult.start.get('minute') || 0;

    // Create new time in user's timezone
    const newStartMoment = moment.tz({ year, month, day, hour, minute }, userTimeZone);
    const newEndMoment = newStartMoment.clone().add(1, 'hour');

    const newStartISO = newStartMoment.format('YYYY-MM-DDTHH:mm:ss.SSSZ');
    const newEndISO = newEndMoment.format('YYYY-MM-DDTHH:mm:ss.SSSZ');

    // Find the meeting to update
    const { findSalesforceMeeting } = require("../salesforce");

    let meeting;
    if (identifierType === 'subject') {
      meeting = await findSalesforceMeeting({ subject: identifier }, teamsChatId);
    } else if (identifierType === 'datetime') {
      // Parse the identifier datetime
      let parsedIdentifierDate = chrono.parseDate(identifier, new Date(), { forwardDate: true });
      if (!parsedIdentifierDate) {
        await context.sendActivity("❌ Could not understand the meeting time to reschedule.");
        return;
      }

      const idChronoResult = chrono.parse(identifier, new Date(), { forwardDate: true })[0];
      const idYear = idChronoResult.start.get('year');
      const idMonth = idChronoResult.start.get('month') - 1;
      const idDay = idChronoResult.start.get('day');
      const idHour = idChronoResult.start.get('hour') || 0;
      const idMinute = idChronoResult.start.get('minute') || 0;

      const identifierMoment = moment.tz({ year: idYear, month: idMonth, day: idDay, hour: idHour, minute: idMinute }, userTimeZone);

      meeting = await findSalesforceMeeting({ dateTime: identifierMoment.format('YYYY-MM-DDTHH:mm:ss.SSSZ') }, teamsChatId);
    }

    if (!meeting) {
      await context.sendActivity(`❌ Could not find a meeting matching "${identifier}". Please check the subject or time and try again.`);
      return;
    }

    // Update the meeting
    const { updateSalesforceMeeting } = require("../salesforce");
    const updateRes = await updateSalesforceMeeting(teamsChatId, meeting.Id, {
      startDateTime: newStartISO,
      endDateTime: newEndISO
    });

    if (updateRes.status === "success") {
      const localDisplay = newStartMoment.format("hh:mm A, DD MMM YYYY");
      await context.sendActivity(
        `✅ Meeting **"${meeting.Subject}"** rescheduled to ${localDisplay}.`
      );
    } else {
      await context.sendActivity(`❌ Failed to reschedule meeting: ${updateRes.message}`);
    }

  } catch (err) {
    console.error("Error updating Salesforce meeting:", err);
    await context.sendActivity(`❌ Error updating meeting: ${err.message}`);
  }
});

// NEW ACTION: Cancel Meeting
app.ai.action("CancelSalesforceMeeting", async (context, state, parameters) => {
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    const { identifier, identifierType } = parameters;

    if (!identifier) {
      await context.sendActivity("❌ Please provide the meeting subject or date/time to cancel.");
      return;
    }

    // Get user's timezone
    const userTimeZone = "Asia/Kolkata";

    // Find the meeting to cancel
    const { findSalesforceMeeting } = require("../salesforce");

    let meeting;
    if (identifierType === 'subject') {
      meeting = await findSalesforceMeeting({ subject: identifier }, teamsChatId);
    } else if (identifierType === 'datetime') {
      // Parse the identifier datetime
      let parsedIdentifierDate = chrono.parseDate(identifier, new Date(), { forwardDate: true });
      if (!parsedIdentifierDate) {
        await context.sendActivity("❌ Could not understand the meeting time to cancel.");
        return;
      }

      const idChronoResult = chrono.parse(identifier, new Date(), { forwardDate: true })[0];
      const idYear = idChronoResult.start.get('year');
      const idMonth = idChronoResult.start.get('month') - 1;
      const idDay = idChronoResult.start.get('day');
      const idHour = idChronoResult.start.get('hour') || 0;
      const idMinute = idChronoResult.start.get('minute') || 0;

      const identifierMoment = moment.tz({ year: idYear, month: idMonth, day: idDay, hour: idHour, minute: idMinute }, userTimeZone);

      meeting = await findSalesforceMeeting({ dateTime: identifierMoment.format('YYYY-MM-DDTHH:mm:ss.SSSZ') }, teamsChatId);
    }

    if (!meeting) {
      await context.sendActivity("❌ Could not find the meeting to cancel.");
      return;
    }

    // Cancel the meeting
    const { cancelSalesforceMeeting } = require("../salesforce");
    const cancelRes = await cancelSalesforceMeeting(teamsChatId, meeting.Id);

    if (cancelRes.status === "success") {
      const meetingTime = moment(meeting.StartDateTime).tz(userTimeZone).format("hh:mm A, DD MMM YYYY");
      await context.sendActivity(
        `✅ Meeting **"${meeting.Subject}"** scheduled for ${meetingTime} has been cancelled.`
      );
    } else {
      await context.sendActivity(`❌ Failed to cancel meeting: ${cancelRes.message}`);
    }

  } catch (err) {
    console.error("Error cancelling Salesforce meeting:", err);
    await context.sendActivity(`❌ Error cancelling meeting: ${err.message}`);
  }
});


app.ai.action("CreateOpportunityFromLatestEmail", async (context, state, parameters) => {
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    
    // Check Outlook authentication
    const token = await getOutlookToken(teamsChatId);
    const isAuthenticated = await initializeConversationStateOutlook(context, state, teamsChatId);
    
    if (!isAuthenticated || !token) {
      const outlookLoginCard = getOutlookLoginCard(context);
      await context.sendActivity({
        attachments: [CardFactory.adaptiveCard(outlookLoginCard)]
      });
      await context.sendActivity(
        MessageFactory.text("🔒 Please authenticate with Outlook using the card above to create an opportunity from your latest email.")
      );
      return "User authentication required - login card sent";
    }

    console.log("Fetching latest email from Outlook...");
    
    // Get the latest email (limit = 1)
    const emailResponse = await getRecentEmails(context, state, 1);
    console.log("Latest email response:", JSON.stringify(emailResponse, null, 2));
    
    if (emailResponse.status !== "success" || !emailResponse.data || emailResponse.data.length === 0) {
      await context.sendActivity(
        MessageFactory.text("📧 No emails found in your inbox.")
      );
      return "No emails found";
    }

    const latestEmail = emailResponse.data[0];
    
    // Generate opportunity details from email
    const opportunityDetails = await generateOpportunityFromEmail(context, state, latestEmail, teamsChatId);
    
    if (!opportunityDetails.success) {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to create opportunity: ${opportunityDetails.message}`)
      );
      return `Failed to create opportunity: ${opportunityDetails.message}`;
    }

    await context.sendActivity(
      MessageFactory.text(
        `✅ **Opportunity Created Successfully from Latest Email!**\n\n` +
        `🆔 **Opportunity ID:** ${opportunityDetails.id}\n` +
        `📋 **Name:** ${opportunityDetails.name}\n` +
        `📊 **Stage:** ${opportunityDetails.stageName}\n` +
        `📅 **Close Date:** ${opportunityDetails.closeDate}\n` +
        `📧 **Source Email:** ${latestEmail.subject || "No subject"} from ${latestEmail.from.emailAddress.name}`
      )
    );
    
    return `Successfully created Salesforce opportunity "${opportunityDetails.name}" from latest email`;

  } catch (error) {
    console.error("Error creating opportunity from latest email:", error);
    const errorMessage = error.response?.data?.error?.message || error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error creating opportunity from latest email: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});


// app.message("/outlook", async (context, state) => {
//   try {

//     // Initilize the conversation state
//     initializeConversationState(state);
//     const userId = context.activity.from.id;
//     const teamsChatId = context.activity.channelData?.teamsChatId || userId;

//     const outlookLoginCard = getOutlookLoginCard(context);

//     await context.sendActivity({
//       attachments: [CardFactory.adaptiveCard(outlookLoginCard)]
//     });
//     return "Outlook login card sended successfully";


//   } catch (error) {
//     console.error("Error in sending outlook message card:", error)
//     return `Error in outlook login card send ${error.message}`
//   }
// })



app.activity(ActivityTypes.Message, async (context, state) => {
  try {
    // initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    console.log("context.activity:", context.activity)
    if (context.activity.text?.startsWith("/")) {
      return;
    }
    const { status, accessToken, refreshToken, instanceUrl } = await getUserToken(teamsChatId, "salesforce");
    if (!status) {
      console.log("User is not authenticated with Salesforce, sending login card.");
      const salesforceLoginCard = await getSalesforceLoginCard(context, userId);
      // Send login card to user
      await context.sendActivity({
        attachments: [CardFactory.adaptiveCard(salesforceLoginCard)]
      });
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
    // await context.sendActivity(
    // MessageFactory.text(
    //   "👋 **Welcome to your AI-powered Salesforce CRM assistant!**\n\n" +
    //   "To get started, we need to connect your SalesForce CRM account. " +
    //   "Once connected, you can ask me questions about your CRM data using natural language!"
    // )
    // );
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




// Handles Logout Functionality
app.ai.action("Logout", async (context, state, parameters) => {
  try {

    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    console.log("Logout action called:", teamsChatId)

    let deletetoken = await deleteUserToken(teamsChatId, "salesforce");
    if (!deletetoken.status) {
      console.log("No token found for user, nothing to delete.");
      return "❌ No Salesforce account connected to log out from.";
    }

    return "✅ Logged out successfully from Salesforce account."

  } catch (error) {
    console.error("Error in Logout:", error);
    return `❌Error in logout: ${error.message || "Unknown error"}.`
  }
})


module.exports = app;