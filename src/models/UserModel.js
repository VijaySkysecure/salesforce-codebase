const mongoose = require("mongoose");


/**
 * @Schema This schema is related to user.
 */
const UserSchema = new mongoose.Schema({
    instanceUrl: {
        type: String,
        trim: true
    },
    teamsChatId: {
        type: String,
        trim: true
    },
    accessToken: {
        type: String
    },
    refreshToken: {
        type: String
    },
    signature:{
        type:String
    },
    instanceUrl:{
        type:String
    }
}, {
    versionKey: false,
    timestamps: true
})


const User = mongoose.model("user", UserSchema);

module.exports = User;