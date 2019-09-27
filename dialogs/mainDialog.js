// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { LogoutDialog } = require('./logoutDialog');
const { OAuthHelpers } = require('../oAuthHelpers');
const { DialogHelper } = require('./dialogHelper');
const { SelectReportDialog } = require('./selectReportDialog');
const { SelectReportResultDialog } = require('./selectReportResultDialog');
const { SelectGlossaryTermDialog } = require('./selectGlossaryTermDialog');
const { SelectGlossaryTermResultDialog } = require('./selectGlossaryTermResultDialog');
const { SearchGlossaryTermDialog } = require('./searchGlossaryTermDialog');

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const OAUTH_PROMPT = 'oAuthPrompt';
const CHOICE_PROMPT = 'choicePrompt';
const TEXT_PROMPT = 'textPrompt';
const SELECT_GLOSSARY_TERM_DIALOG = 'selectGlossaryTermDialog';
const SEARCH_GLOSSARY_TERM_DIALOG = 'searchGlossaryTermDialog';

const WelcomeCard = require('../bots/resources/welcomeCard.json');

class MainDialog extends LogoutDialog {
    constructor() {
        super('MainDialog');

        this.dialogHelper = new DialogHelper();
        this.selectReportDialog = new SelectReportDialog();
        this.selectReportResultDialog = new SelectReportResultDialog();
        this.selectGlossaryTermDialog = new SelectGlossaryTermDialog();
        this.selectGlossaryTermResultDialog = new SelectGlossaryTermResultDialog();


        this.addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new OAuthPrompt(OAUTH_PROMPT, {
                connectionName: process.env.ConnectionName,
                text: 'Please log in with your credentials and enter the validation code into this chat window to complete the log in process',
                title: 'Log In with My Credentials',
                timeout: 300000
            }))
            .addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new SearchGlossaryTermDialog(SEARCH_GLOSSARY_TERM_DIALOG))
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
                this.promptStep.bind(this),
                this.loginStep.bind(this),
                this.commandStep.bind(this),
                this.processStep.bind(this)
            ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;

    }

    /**
     * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} turnContext
     * @param {*} accessor
     */
    async run(turnContext, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async promptStep(step) {

      switch (step.context.activity.text) {
        case 'Log In As Guest':
            console.log(step.context.activity.text)
            await step.endDialog();
            return await this.selectReportDialog.destinationStep(step);
            break;
        default:
            return await step.beginDialog(OAUTH_PROMPT);
      }

    }

    async loginStep(step) {
        // Get the token from the previous step. Note that we could also have gotten the
        // token directly from the prompt itself. There is an example of this in the next method.
        const tokenResponse = step.result;

        if (tokenResponse) {

          if (step.context.activity.value){

            if (step.context.activity.value.action === 'report_name_selector_value'){

              return await this.selectReportResultDialog.onTurn(step.context);
            }

            if (step.context.activity.value.action === 'glossary_term_selector_value'){

              return await this.selectGlossaryTermResultDialog.onTurn(step, step.context, tokenResponse);
            }

              //console.log(step.context.activity.value.action)
              return await step.endDialog();
          }else{

            switch (step.context.activity.text) {
            case 'Single-View Glossary':
                return await this.selectGlossaryTermDialog.destinationStep(step, tokenResponse, step.context.activity.text);
                break;
            case 'Multi-View Glossary':
                return await this.selectGlossaryTermDialog.destinationStep(step, tokenResponse, step.context.activity.text);
                break;
            case 'Search the Glossary':
                //return await this.selectGlossaryTermDialog.searchStep(step, tokenResponse);
                return await step.beginDialog(SEARCH_GLOSSARY_TERM_DIALOG);
                break;
            default:
                //await step.context.sendActivity(`Your token is ${ tokenResponse.token }`);

                await step.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What would you like to do?','')] });

                return await step.prompt(CHOICE_PROMPT, {
                    prompt: '',
                    choices: ChoiceFactory.toChoices(['Who Am I?', 'Glossary', 'Archer Reports'])
                });

            }

          }



        }else{

          await step.context.sendActivity('Login was not successful please try again.');
          return await step.endDialog();

        }

    }

    async commandStep(step) {
        step.values['command'] = step.result;

        // Call the prompt again because we need the token. The reasons for this are:
        // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
        // about refreshing it. We can always just call the prompt again to get the token.
        // 2. We never know how long it will take a user to respond. By the time the
        // user responds the token may have expired. The user would then be prompted to login again.
        //
        // There is no reason to store the token locally in the bot because we can always just call
        // the OAuth prompt to get the token or get a new token if needed.
        return await step.beginDialog(OAUTH_PROMPT);
    }

    async processStep(step) {
        if (step.result) {
            // We do not need to store the token in the bot. When we need the token we can
            // send another prompt. If the token is valid the user will not need to log back in.
            // The token will be available in the Result property of the task.
            const tokenResponse = step.result;

            // If we have the token use the user is authenticated so we may use it to make API calls.
            if (tokenResponse && tokenResponse.token) {

                //console.log(step.values.command)

                //const parts = (step.values['command'] || '').toLowerCase().split(' ');

                //const command = parts[0];

                switch (step.values.command.value) {
                case 'Who Am I?':
                    await OAuthHelpers.listMe(step.context, tokenResponse);
                    break;
                case 'Glossary':
                    //await this.selectGlossaryTermDialog.filterStep(step);
                    //await step.beginDialog(SELECT_GLOSSARY_TERM_DIALOG);

                    await step.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Single-View Glossary, Multi-View Glossary or Search the Glossary?','')] });

                    await step.prompt(CHOICE_PROMPT, {
                        prompt: '',
                        choices: ChoiceFactory.toChoices(['Single-View Glossary', 'Multi-View Glossary', 'Search the Glossary'])
                    });

                    break;
                case 'Archer Reports':
                    await this.selectReportDialog.destinationStep(step);
                    break;
                default:
                    //await step.context.sendActivity(`Your token is ${ tokenResponse.token }`);
                }
            }
        } else {
            await step.context.sendActivity('We couldn\'t log you in. Please try again later.');
        }

        return await step.endDialog();
    }
}

module.exports.MainDialog = MainDialog;
