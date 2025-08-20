const config = require("../config");
const {TurnContext } = require("botbuilder");
// Helper: Generate Salesforce login Adaptive Card
function getSalesforceLoginCard(context) {
    const { salesforceClientId, salesforceRedirectUri, salesforceOauthScope} = config;
      // Create the auth state with userId and nonce
      const userId = context.activity.from.id;
      const conversationReference = TurnContext.getConversationReference(context.activity);
      const nonce = Math.random().toString(36).substring(2);
      const authState = Buffer.from(JSON.stringify({ userId, nonce, conversationReference })).toString("base64url");
      console.log(salesforceClientId, salesforceRedirectUri, salesforceOauthScope)


    return {
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        type: "AdaptiveCard",
        version: "1.5",
        body: [
            {
                type: "Container",
                style: "emphasis",
                bleed: true,
                items: [
                    {
                        type: "Image",
                        url: "https://logos-world.net/wp-content/uploads/2020/10/Salesforce-Logo.png",
                        horizontalAlignment: "Center",
                        size: "Large",
                        spacing: "None"
                    },
                    {
                        type: "TextBlock",
                        text: "⚡ Connect to Salesforce CRM",
                        wrap: true,
                        weight: "Bolder",
                        size: "ExtraLarge",
                        horizontalAlignment: "Center",
                        color: "Accent",
                        spacing: "Medium"
                    },
                    {
                        type: "FactSet",
                        facts: [
                            { title: "🔒 Security:", value: "OAuth 2.0 Protected" },
                            { title: "🚀 Access:", value: "Real-time CRM Data" },
                            { title: "🤖 AI Power:", value: "Smart Query & Analytics" }
                        ],
                        spacing: "Large"
                    }
                ]
            }
        ],
        actions: [
            {
                type: "Action.OpenUrl",
                title: "🚀 Login to Salesforce",
                url: `https://login.salesforce.com/services/oauth2/authorize?response_type=code&client_id=${salesforceClientId}&redirect_uri=${salesforceRedirectUri}&state=${authState}`,
                style: "positive"
            },
            {
                type: "Action.OpenUrl",
                title: "🔎 Learn More",
                url: "https://www.salesforce.com/",
                style: "default"
            }
        ]
    };
}

module.exports = {
    getSalesforceLoginCard
};
