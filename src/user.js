const User = require("./models/UserModel");

const storeUserToken = async (teamsChatId, type, tokenResponse) => {
    try {

        const record = {
            type,
            teamsChatId,
            accessToken: tokenResponse.access_token,
            instanceUrl: tokenResponse?.instance_url,
            signature: tokenResponse?.signature,
        };

        if (tokenResponse.refresh_token) {
            record.refreshToken = tokenResponse.refresh_token;
        }

        const user = await User.findOneAndUpdate(
            { teamsChatId, type },
            record,
            { new: true, upsert: true }
        )
    } catch (error) {

    }
}

async function getUserToken(teamsChatId, type) {
    try {
        const user = await User.findOne({ teamsChatId, type })
        if (user) {
            return {
                status: true,
                accessToken: user.accessToken,
                refreshToken: user.refreshToken,
                instanceUrl: user?.instanceUrl ?? "",
            };
        }
        return {
            status: false
        };
    } catch (error) {
        console.error("Error in getting user token:", error)
        return { status: false }
    }
}

module.exports = {
    storeUserToken,
    getUserToken
}