const config = require("../config");
const { TurnContext } = require("botbuilder");
// Helper: Generate Salesforce login Adaptive Card
function getOutlookLoginCard(context) {
    const { microsoftClientId, microsoftRedirectUrl, microsoftOauthScope } = config;
    // Create the auth state with userId and nonce
    const userId = context.activity.from.id;
    const conversationReference = TurnContext.getConversationReference(context.activity);
    const nonce = Math.random().toString(36).substring(2);
    const authState = Buffer.from(JSON.stringify({ userId, nonce, conversationReference })).toString("base64url");
    console.log(microsoftClientId,microsoftRedirectUrl , microsoftOauthScope)


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
                        url: "https://img.icons8.com/color/512/microsoft-outlook-2019.png",
                        horizontalAlignment: "Center",
                        size: "Large",
                        spacing: "None"
                    },
                    {
                        type: "TextBlock",
                        text: "⚡ Connect to Outlook Calendar",
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
                title: "🚀 Login to Outlook",
                url: `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${microsoftClientId}&response_type=code&redirect_uri=${microsoftRedirectUrl}&response_mode=query&scope=${microsoftOauthScope}&state=${authState}&prompt=select_account`,
                style: "positive"
            },
            {
                type: "Action.OpenUrl",
                title: "🔎 Learn More",
                url: "https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow",
                style: "default"
            }
        ]
    };
}

module.exports = {
    getOutlookLoginCard
};
