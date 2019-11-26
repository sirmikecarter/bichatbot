// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { CardFactory } = require('botbuilder-core');
const { DialogBot } = require('./dialogBot');
const WelcomeCard = require('./resources/welcomeCard.json');

class AuthBot extends DialogBot {
    constructor(conversationState, userState, dialog) {
        super(conversationState, userState, dialog);

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                // if (membersAdded[cnt].id !== context.activity.recipient.id) {
                //     await context.sendActivity('Welcome to Authentication Bot on MSGraph. Type anything to get logged in. Type \'logout\' to sign-out.');
                // }

                if (membersAdded[cnt].id === context.activity.recipient.id) {
                    //await context.sendActivity('Welcome to Authentication Bot on MSGraph. Type anything to get logged in. Type \'logout\' to sign-out.');
                    const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                    await context.sendActivity({ attachments: [welcomeCard] });
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onTokenResponseEvent(async (context, next) => {
            console.log('Running dialog with Token Response Event Activity.');

            // Run the Dialog with the new Token Response Event Activity.
            await this.dialog.run(context, this.dialogState);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.AuthBot = AuthBot;
