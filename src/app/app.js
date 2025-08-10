const { MemoryStorage, MessageFactory } = require("botbuilder");
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
    enable_feedback_loop: true,
  },
});

app.feedbackLoop(async (context, state, feedbackLoopData) => {
  //add custom feedback process logic here
  console.log("Your feedback is " + JSON.stringify(context.activity.value));
});

const {
  createSalesforceLead,
  // getSalesforceLeads,
  updateSalesforceLead,
  deleteSalesforceLead,
} = require("../salesforce");


app.ai.action("GetSalesforceLeads", async (context, state, parameters) => {
  console.log("GetSalesforceLeads action called with parameters:", parameters);
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;
    
    const limit = Math.min(parameters.limit || 20, 200);
    const query = `SELECT Id, FirstName, LastName, Company, Status, Email, Phone FROM Lead ORDER BY CreatedDate DESC LIMIT ${limit}`;

    const config = require("../config");
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
      await context.sendActivity(MessageFactory.text("📊 No leads found in your Salesforce CRM."));
      return "No leads found";
    }

    const formattedLeads = records.map((l) => ({
      id: l.Id,
      name: `${l.FirstName || ''} ${l.LastName || ''}`.trim() || "—",
      company: l.Company || "—",
      status: l.Status || "—",
      email: l.Email || "—",
      phone: l.Phone || "—",
    }));

    state.conversation.formattedLeads = JSON.stringify(formattedLeads, null, 2);
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
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    if (!parameters.leadId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Lead ID to update a lead.")
      );
      return "Missing required parameters";
    }

    const updateFields = {
      ...(parameters.firstName && { FirstName: parameters.firstName }),
      ...(parameters.lastName && { LastName: parameters.lastName }),
      ...(parameters.company && { Company: parameters.company }),
      ...(parameters.email && { Email: parameters.email }),
      ...(parameters.phone && { Phone: parameters.phone }),
      ...(parameters.status && { Status: parameters.status }),
      ...(parameters.title && { Title: parameters.title }),
      ...(parameters.leadSource && { LeadSource: parameters.leadSource }),
      ...(parameters.industry && { Industry: parameters.industry }),
    };

    if (Object.keys(updateFields).length === 0) {
      await context.sendActivity(
        MessageFactory.text("❌ No fields provided to update. Please include at least one field like firstName, email, etc.")
      );
      return "No fields provided to update";
    }

    console.log(`Updating lead ${parameters.leadId} in Salesforce CRM...`);
    const { updateSalesforceLeadById } = require("../salesforce");

    const response = await updateSalesforceLead(context, state, parameters.leadId, updateFields);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(
          `✅ **Lead Updated Successfully!**\n\n` +
          `🆔 **Lead ID:** ${parameters.leadId}\n` +
          (parameters.firstName ? `👤 **First Name:** ${parameters.firstName}\n` : "") +
          (parameters.lastName ? `👤 **Last Name:** ${parameters.lastName}\n` : "") +
          (parameters.company ? `🏢 **Company:** ${parameters.company}\n` : "") +
          (parameters.email ? `📧 **Email:** ${parameters.email}\n` : "") +
          (parameters.phone ? `📱 **Phone:** ${parameters.phone}\n` : "") +
          (parameters.status ? `📊 **Status:** ${parameters.status}\n` : "") +
          (parameters.title ? `💼 **Title:** ${parameters.title}\n` : "") +
          (parameters.leadSource ? `🌐 **Source:** ${parameters.leadSource}\n` : "") +
          (parameters.industry ? `🏭 **Industry:** ${parameters.industry}\n` : "")
        )
      );
      return `Successfully updated Salesforce lead ${parameters.leadId}`;
    } else {
      await context.sendActivity(
        MessageFactory.text(`❌ Failed to update lead: ${response.message || "Unknown error"}.`)
      );
      return `Failed to update lead: ${response.message || "Unknown error"}`;
    }
  } catch (error) {
    console.error("Error updating Salesforce lead:", error);
    const errorMessage = error.message || "Unknown error";
    await context.sendActivity(
      MessageFactory.text(`❌ Error updating lead: ${errorMessage}. Please try again.`)
    );
    return `Error occurred: ${errorMessage}`;
  }
});



app.ai.action("DeleteSalesforceLead", async (context, state, parameters) => {
  try {
    initializeConversationState(state);
    const userId = context.activity.from.id;
    const teamsChatId = context.activity.channelData?.teamsChatId || userId;

    if (!parameters.leadId) {
      await context.sendActivity(
        MessageFactory.text("❌ Missing required information. I need at least: Lead ID to delete a lead.")
      );
      return "Missing required parameters";
    }

    console.log(`Deleting lead ${parameters.leadId} from Salesforce CRM...`);
    const { deleteSalesforceLeadById } = require("../salesforce");

    const response = await deleteSalesforceLead(context, state, parameters.leadId);

    if (response.status === "success") {
      await context.sendActivity(
        MessageFactory.text(`✅ **Lead Deleted Successfully!**\n\n🆔 **Lead ID:** ${parameters.leadId}`)
      );
      return `Successfully deleted lead ${parameters.leadId}`;
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




module.exports = app;
