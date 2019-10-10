// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder');
const { ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');
const { SimpleGraphClient } = require('./simple-graph-client');
const { DialogHelper } = require('./dialogs/dialogHelper');

const CHOICE_PROMPT = 'choicePrompt';

/**
 * These methods call the Microsoft Graph API. The following OAuth scopes are used:
 * 'OpenId' 'email' 'Mail.Send.Shared' 'Mail.Read' 'profile' 'User.Read' 'User.ReadBasic.All'
 * for more information about scopes see:
 * https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference
 */
class OAuthHelpers {
    /**
     * Enable the user to send an email via the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     * @param {string} emailAddress The email address of the recipient.
     */

    static async sendMail(context, tokenResponse, emailAddress) {
        if (!context) {
            throw new Error('OAuthHelpers.sendMail(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.sendMail(): `tokenResponse` cannot be undefined.');
        }

        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();

        await client.sendMail(
            emailAddress,
            `Message from a bot!`,
            `Hi there! I had this message sent from a bot. - Your friend, ${ me.displayName }`
        );
        await context.sendActivity(`I sent a message to ${ emailAddress } from your account.`);
    }

    /**
     * Displays information about the user in the bot.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     */
    static async listMe(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.listMe(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listMe(): `tokenResponse` cannot be undefined.');
        }

        // Pull in the data from Microsoft Graph.
        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();

        this.dialogHelper = new DialogHelper();

        // await context.sendActivity(`You are: ${ me.displayName }`);
        // await context.sendActivity(`You're department is: ${ me.jobTitle }`);

        if (me.jobTitle === 'FINO'){
          await context.sendActivity({ attachments: [this.dialogHelper.createUserCard('https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTRHuG7ELWmAZNaE11lKdHTMSOGVb1xnQjh_KWz0I9VHAOk31HLpw', me.displayName, me.jobTitle)] });
          // await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Youre 16 years away from Retiring !','')] });

        }
        //
         if (me.jobTitle === 'ACTO'){
           await context.sendActivity({ attachments: [this.dialogHelper.createUserCard('https://www2.deloitte.com/content/dam/Deloitte/us/Images/promo_images/profile-pictures/us-emily-mossburg.jpg/_jcr_content/renditions/cq5dam.web.250.250.mobile.jpeg', me.displayName, me.jobTitle)] });
           await context.sendActivity({ attachments: [this.dialogHelper.createBotCard('You are 3 years away from Retiring !!!','')] });
         }

      //  await context.sendActivity({ attachments: [this.dialogHelper.createUserCard(me.displayName, 'Financial Specialist', me.jobTitle)] });



        var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
        await context.sendActivity(reply);


    }

    /**
     * Lists the user's collected email.
     * @param {TurnContext} context A TurnContext instance containing all the data needed for processing this conversation turn.
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     */
    static async listRecentMail(context, tokenResponse) {
        if (!context) {
            throw new Error('OAuthHelpers.listRecentMail(): `context` cannot be undefined.');
        }
        if (!tokenResponse) {
            throw new Error('OAuthHelpers.listRecentMail(): `tokenResponse` cannot be undefined.');
        }

        var client = new SimpleGraphClient(tokenResponse.token);
        var messages = await client.getRecentMail();
        if (Array.isArray(messages)) {
            let numberOfMessages = 0;
            if (messages.length > 5) {
                numberOfMessages = 5;
            }

            const reply = { attachments: [], attachmentLayout: AttachmentLayoutTypes.Carousel };
            for (let cnt = 0; cnt < numberOfMessages; cnt++) {
                const mail = messages[cnt];
                const card = CardFactory.heroCard(
                    mail.subject,
                    mail.bodyPreview,
                    [{ alt: 'Outlook Logo', url: 'https://botframeworksamples.blob.core.windows.net/samples/OutlookLogo.jpg' }],
                    [],
                    { subtitle: `${ mail.from.emailAddress.name } <${ mail.from.emailAddress.address }>` }
                );
                reply.attachments.push(card);
            }
            await context.sendActivity(reply);
        } else {
            await context.sendActivity('Unable to find any recent unread mail.');
        }
    }
}

exports.OAuthHelpers = OAuthHelpers;
