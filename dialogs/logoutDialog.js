// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer, QnAMaker } = require('botbuilder-ai');

const { ActivityTypes } = require('botbuilder');
const { ComponentDialog } = require('botbuilder-dialogs');
const { DialogHelper } = require('./dialogHelper');


const CHOICE_PROMPT = 'choicePrompt';
const QNA_TOP_N = 1;
const QNA_CONFIDENCE_THRESHOLD = 0.5;

class LogoutDialog extends ComponentDialog {

  constructor(id) {
      super(id || 'logoutDialog');

      const luisApplication = {
          applicationId: process.env.LuisAppId,
          azureRegion: process.env.LuisAPIHostName,
          // CAUTION: Authoring key is used in this example as it is appropriate for prototyping.
          // When implimenting for deployment/production, assign and use a subscription key instead of an authoring key.
          endpointKey: process.env.LuisAPIKey
      };

      const luisPredictionOptions = {
          spellCheck: true,
          bingSpellCheckSubscriptionKey: process.env.BingSpellCheck

      };

      this.qnaRecognizer = new QnAMaker({
          knowledgeBaseId: process.env.QnAKbId,
          endpointKey: process.env.QnAEndpointKey,
          host: process.env.QnAHostname
      });

      this.luisRecognizer = new LuisRecognizer(luisApplication, luisPredictionOptions);

  }

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

                await innerDc.prompt(CHOICE_PROMPT, {
                    prompt: '',
                    choices: ChoiceFactory.toChoices(['Log In'])
                });

                //await innerDc.context.sendActivity('You have been signed out.');
                return await innerDc.cancelAllDialogs();
            }

            if (text === 'menu') {
              console.log(text)

            }

            const dispatchResults = await this.luisRecognizer.recognize(innerDc.context);
            const dispatchTopIntent = LuisRecognizer.topIntent(dispatchResults);

            //console.log(dispatchTopIntent)

          //  console.log(dispatchResults.text)


            switch (dispatchTopIntent) {
              case 'General':
                //console.log(dispatchResults.intents)
                //console.log(dispatchTopIntent)
                const qnaResult = await this.qnaRecognizer.generateAnswer(dispatchResults.text, QNA_TOP_N, QNA_CONFIDENCE_THRESHOLD);
                if (!qnaResult || qnaResult.length === 0 || !qnaResult[0].answer){
                  //await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                }else{
                  await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                }
                //console.log(qnaResult[0].answer)
                break;
              case 'Glossary':
                //console.log(dispatchResults.intents)
                console.log(dispatchTopIntent)
                break;
              case 'Reports':
                  //console.log(dispatchResults.intents)
                console.log(dispatchTopIntent)
                break;
              case 'Weather':
                //console.log(dispatchResults.intents)
                console.log(dispatchTopIntent)
                break;
              case 'Sports':
                //console.log(dispatchResults.intents)
                console.log(dispatchTopIntent)
                break;

              default:
                //console.log(dispatchResults.intents);
                break;
            }


        }
    }
}

module.exports.LogoutDialog = LogoutDialog;
