// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { MenuDialog } = require('./menuDialog');
const { OAuthHelpers } = require('./helpers/oAuthHelpers');
const { DialogHelper } = require('./helpers/dialogHelper');
const { SearchReportDialog } = require('./searchReportDialog');
const { SelectReportDialog } = require('./selectReportDialog');
const { SelectReportResultDialog } = require('./selectReportResultDialog');
const { SelectGlossaryTermDialog } = require('./selectGlossaryTermDialog');
const { SelectGlossaryTermResultDialog } = require('./selectGlossaryTermResultDialog');
const { SearchGlossaryTermDialog } = require('./searchGlossaryTermDialog');
const { GuestLogInDialog } = require('./guestLogInDialog');

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const OAUTH_PROMPT = 'oAuthPrompt';
const CHOICE_PROMPT = 'choicePrompt';
const TEXT_PROMPT = 'textPrompt';
const SEARCH_GLOSSARY_TERM_DIALOG = 'searchGlossaryTermDialog';
const SELECT_GLOSSARY_TERM_DIALOG = 'selectGlossaryTermDialog';
const SELECT_GLOSSARY_TERM_RESULT_DIALOG = 'selectGlossaryTermResultDialog'
const SEARCH_REPORT_DIALOG = 'searchReportDialog';
const SELECT_REPORT_DIALOG = 'selectReportDialog';
const SELECT_REPORT_RESULT_DIALOG = 'selectReportResultDialog';
const GUEST_LOG_IN_DIALOG = 'guestLogInDialog';

const WelcomeCard = require('../bots/resources/welcomeCard.json');

class MainDialog extends MenuDialog {
    constructor() {
        super('MainDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          statusUpdate: false,
        };

        this.addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new OAuthPrompt(OAUTH_PROMPT, {
                connectionName: process.env.ConnectionName,
                text: 'Please log in with your credentials and enter the validation code into this chat window to complete the log in process',
                title: 'Log In with My Credentials',
                timeout: 300000
            }))
            .addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new SearchGlossaryTermDialog(SEARCH_GLOSSARY_TERM_DIALOG))
            .addDialog(new SelectGlossaryTermDialog(SELECT_GLOSSARY_TERM_DIALOG))
            .addDialog(new SelectGlossaryTermResultDialog(SELECT_GLOSSARY_TERM_RESULT_DIALOG))
            .addDialog(new SearchReportDialog(SEARCH_REPORT_DIALOG))
            .addDialog(new SelectReportDialog(SELECT_REPORT_DIALOG))
            .addDialog(new SelectReportResultDialog(SELECT_REPORT_RESULT_DIALOG))
            .addDialog(new GuestLogInDialog(GUEST_LOG_IN_DIALOG))
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

    async promptStep(stepContext) {

      switch (stepContext.context.activity.text) {
        case 'Log In As Guest':
            //console.log(stepContext.context.activity.text)
            return await stepContext.beginDialog(GUEST_LOG_IN_DIALOG);
            break;
        default:
            //return await stepContext.beginDialog(OAUTH_PROMPT);
            return await stepContext.endDialog();
      }

    }

    async loginStep(stepContext) {
        // Get the token from the previous step. Note that we could also have gotten the
        // token directly from the prompt itself. There is an example of this in the next method.
        const tokenResponse = stepContext.result;

        if (tokenResponse) {

          if (stepContext.context.activity.value){

            if (stepContext.context.activity.value.action === 'report_name_selector_value'){
              return await stepContext.beginDialog(SELECT_REPORT_RESULT_DIALOG);
            }

            if (stepContext.context.activity.value.action === 'glossary_term_selector_value'){
              return await stepContext.beginDialog(SELECT_GLOSSARY_TERM_RESULT_DIALOG, { tokenResponse: tokenResponse});
            }

            return await stepContext.endDialog();
          }else{

            switch (stepContext.context.activity.text) {
            case 'Select a Term':
                return await stepContext.beginDialog(SELECT_GLOSSARY_TERM_DIALOG, { tokenResponse: tokenResponse});
                break;
            case 'See All Terms':
                return await stepContext.beginDialog(SELECT_GLOSSARY_TERM_DIALOG, { tokenResponse: tokenResponse});
                break;
            case 'Glossary Search':
                return await stepContext.beginDialog(SEARCH_GLOSSARY_TERM_DIALOG);
                break;
            case 'Select a Report':
                return await stepContext.beginDialog(SELECT_REPORT_DIALOG);
                break;
            case 'Search Reports':
                return await stepContext.beginDialog(SEARCH_REPORT_DIALOG);
                break;
            default:
                await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('What would you like to do?','')] });

                return await stepContext.prompt(CHOICE_PROMPT, {
                    prompt: '',
                    choices: ChoiceFactory.toChoices(['Who Am I?', 'Glossary', 'Cognos Reports'])
                });
            }
          }
        }else{
          await stepContext.context.sendActivity('Login was not successful please try again.');
          return await stepContext.endDialog();
        }
    }

    async commandStep(stepContext) {
        stepContext.values['command'] = stepContext.result;
        // Call the prompt again because we need the token. The reasons for this are:
        // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
        // about refreshing it. We can always just call the prompt again to get the token.
        // 2. We never know how long it will take a user to respond. By the time the
        // user responds the token may have expired. The user would then be prompted to login again.
        //
        // There is no reason to store the token locally in the bot because we can always just call
        // the OAuth prompt to get the token or get a new token if needed.
        return await stepContext.beginDialog(OAUTH_PROMPT);
    }

    async processStep(stepContext) {
        if (stepContext.result) {
            // We do not need to store the token in the bot. When we need the token we can
            // send another prompt. If the token is valid the user will not need to log back in.
            // The token will be available in the Result property of the task.
            const tokenResponse = stepContext.result;

            // If we have the token use the user is authenticated so we may use it to make API calls.
            if (tokenResponse && tokenResponse.token) {

                switch (stepContext.values.command.value) {

                case 'Who Am I?':
                    await OAuthHelpers.listMe(stepContext.context, tokenResponse);
                    break;
                case 'Glossary':
                    await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Select a Term in the Glossary, See All Terms in the Glossary or Search the Glossary?','')] });
                    await stepContext.prompt(CHOICE_PROMPT, {
                        prompt: '',
                        choices: ChoiceFactory.toChoices(['Select a Term', 'See All Terms', 'Glossary Search'])
                    });
                    break;
                case 'Cognos Reports':
                    await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Select a Report or Search Reports?','')] });
                    await stepContext.prompt(CHOICE_PROMPT, {
                        prompt: '',
                        choices: ChoiceFactory.toChoices(['Select a Report', 'Search Reports'])
                    });
                    break;
                default:
                    //await stepContext.context.sendActivity(`Your token is ${ tokenResponse.token }`);
                }
            }
        } else {
            await stepContext.context.sendActivity('We couldn\'t log you in. Please try again later.');
        }

        return await stepContext.endDialog();
    }
}

module.exports.MainDialog = MainDialog;
