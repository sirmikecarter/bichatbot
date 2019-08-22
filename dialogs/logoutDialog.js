// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');

const { ActivityTypes } = require('botbuilder');
const { ComponentDialog } = require('botbuilder-dialogs');
const { DialogHelper } = require('./dialogHelper');

const CHOICE_PROMPT = 'choicePrompt';

class LogoutDialog extends ComponentDialog {
    async onBeginDialog(innerDc, options) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onBeginDialog(innerDc, options);
    }

    async onContinueDialog(innerDc) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onContinueDialog(innerDc);
    }

    async interrupt(innerDc) {
        if (innerDc.context.activity.type === ActivityTypes.Message) {
            const text = innerDc.context.activity.text ? innerDc.context.activity.text.toLowerCase() : '';
            if (text === 'logout') {
                // The bot adapter encapsulates the authentication processes.
                const botAdapter = innerDc.context.adapter;
                await botAdapter.signOutUser(innerDc.context, process.env.ConnectionName);

                await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('You have been signed out.','')] });

                return await innerDc.prompt(CHOICE_PROMPT, {
                    prompt: '',
                    choices: ChoiceFactory.toChoices(['Log In'])
                });

                //await innerDc.context.sendActivity('You have been signed out.');
                return await innerDc.cancelAllDialogs();
            }

            if (text === 'menu') {
              console.log(text)

            }
        }
    }
}

module.exports.LogoutDialog = LogoutDialog;
