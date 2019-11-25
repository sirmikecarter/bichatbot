// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { ComponentDialog, ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer, QnAMaker } = require('botbuilder-ai');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { ActivityTypes } = require('botbuilder');
const { DialogHelper } = require('./dialogHelper');
const { GuestLogInDialog } = require('./guestLogInDialog');
const { SearchGlossaryTermDialog } = require('./searchGlossaryTermDialog');
const { SearchReportDialog } = require('./searchReportDialog');
const { SelectSportsTeamDialog } = require('./selectSportsTeamDialog');
const { SearchSoftwareApprovedDialog } = require('./searchSoftwareApprovedDialog');
const { SearchSoftwareInstalledDialog } = require('./searchSoftwareInstalledDialog');
const { SearchSoftwareRAWDialog } = require('./searchSoftwareRAWDialog');
const { SearchSoftwareFinancialDialog } = require('./searchSoftwareFinancialDialog');
const { SearchSportsTeamDialog } = require('./searchSportsTeamDialog');
const { SearchWeatherDialog } = require('./searchWeatherDialog');
const { SearchGlossaryTermTextDialog } = require('./searchGlossaryTermTextDialog');
const { SearchReportTextDialog } = require('./searchReportTextDialog');
const axios = require('axios');
var arraySort = require('array-sort');


const CHOICE_PROMPT = 'choicePrompt';
const OAUTH_PROMPT = 'oAuthPrompt';
const GUEST_LOG_IN_DIALOG = 'guestLogInDialog';
const SELECT_SPORTS_TEAM_DIALOG = 'selectSportsTeamDialog';
const SEARCH_SOFTWARE_APPROVED_DIALOG = 'searchSoftwareApprovedDialog';
const SEARCH_SOFTWARE_INSTALLED_DIALOG = 'searchSoftwareInstalledDialog';
const SEARCH_SOFTWARE_RAW_DIALOG = 'searchSoftwareRAWDialog';
const SEARCH_SOFTWARE_FINANCIAL_DIALOG = 'searchSoftwareFinancialDialog';
const SEARCH_SPORTS_TEAM_DIALOG = 'searchSportsTeamDialog';
const SEARCH_WEATHER_DIALOG = 'searchWeatherDialog';
const SEARCH_GLOSSARY_TERM_TEXT_DIALOG = 'searchGlossaryTermTextDialog';
const SEARCH_REPORT_TEXT_DIALOG = 'searchReportTextDialog';
const SEARCH_REPORT_DIALOG = 'searchReportDialog';
const QNA_TOP_N = 1;
const QNA_CONFIDENCE_THRESHOLD = 0.5;


class LogoutDialog extends ComponentDialog {

  constructor(id) {
      super(id || 'logoutDialog');

      this.state = {

      };

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

      this.addDialog(new SelectSportsTeamDialog(SELECT_SPORTS_TEAM_DIALOG))
      this.addDialog(new SearchSoftwareApprovedDialog(SEARCH_SOFTWARE_APPROVED_DIALOG))
      this.addDialog(new SearchSoftwareInstalledDialog(SEARCH_SOFTWARE_INSTALLED_DIALOG))
      this.addDialog(new SearchSoftwareRAWDialog(SEARCH_SOFTWARE_RAW_DIALOG))
      this.addDialog(new SearchSoftwareFinancialDialog(SEARCH_SOFTWARE_FINANCIAL_DIALOG))
      this.addDialog(new SearchSportsTeamDialog(SEARCH_SPORTS_TEAM_DIALOG))
      this.addDialog(new SearchWeatherDialog(SEARCH_WEATHER_DIALOG))
      this.addDialog(new SearchGlossaryTermTextDialog(SEARCH_GLOSSARY_TERM_TEXT_DIALOG))
      this.addDialog(new SearchReportTextDialog(SEARCH_REPORT_TEXT_DIALOG))

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

            const dispatchResults = await this.luisRecognizer.recognize(innerDc.context);
            const dispatchTopIntent = LuisRecognizer.topIntent(dispatchResults);

            console.log(dispatchTopIntent)
            console.log(dispatchResults.text)


            switch (dispatchTopIntent) {
              case 'General':
                  const qnaResult = await this.qnaRecognizer.generateAnswer(dispatchResults.text, QNA_TOP_N, QNA_CONFIDENCE_THRESHOLD);
                  if (!qnaResult || qnaResult.length === 0 || !qnaResult[0].answer){
                    //await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                  }else{
                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                  }

                  break;
              case 'Glossary':
                  const glossaryTerm = dispatchResults.entities.Term[0];

                  return await innerDc.beginDialog(SEARCH_GLOSSARY_TERM_TEXT_DIALOG, { glossary_term: glossaryTerm});
                  break;
              case 'Reports':

                  const reportApprover = dispatchResults.entities.Report_Approver;
                  const reportClassification = dispatchResults.entities.Report_Classification;
                  const reportDescription = dispatchResults.entities.Report_Description;
                  const reportDesignee = dispatchResults.entities.Report_Designee;
                  const reportDivision = dispatchResults.entities.Report_Division;
                  const reportName = dispatchResults.entities.Report_Name;
                  const reportOwner = dispatchResults.entities.Report_Owner;

                  return await innerDc.beginDialog(SEARCH_REPORT_TEXT_DIALOG, { report_approver: reportApprover, report_classification: reportClassification, report_description: reportDescription, report_designee: reportDesignee, report_division: reportDivision, report_name: reportName,  report_owner:  reportOwner});

                break;
              case 'Weather':

                const cityName = dispatchResults.entities.Cities[0];

                return await innerDc.beginDialog(SEARCH_WEATHER_DIALOG, { city_name: cityName});
                break;
              case 'Sports':

                const teamName = dispatchResults.entities.Sports_Teams[0].toLowerCase();

                return await innerDc.beginDialog(SEARCH_SPORTS_TEAM_DIALOG, { sports_team: teamName});
                break;
              case 'Log_In_As_Guest':

                  return await innerDc.beginDialog(GUEST_LOG_IN_DIALOG);
                  break;
              case 'Software_Approved':

                const searchApprovedTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

                return await innerDc.beginDialog(SEARCH_SOFTWARE_APPROVED_DIALOG, { software_name: searchApprovedTerm});
                break;

              case 'Software_Installed':

                const searchInstalledTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

                return await innerDc.beginDialog(SEARCH_SOFTWARE_INSTALLED_DIALOG, { software_name: searchInstalledTerm});
                break;

              case 'Software_RAW':

                const searchRAWTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

                return await innerDc.beginDialog(SEARCH_SOFTWARE_RAW_DIALOG, { software_name: searchRAWTerm});
                break;

              case 'Software_Financials':

                const searchFinancialTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

                return await innerDc.beginDialog(SEARCH_SOFTWARE_FINANCIAL_DIALOG, { software_name: searchFinancialTerm});
                break;

              case 'Software_All':

                const searchAllTerm = dispatchResults.entities.Software_Name[0].toLowerCase();

                await innerDc.beginDialog(SEARCH_SOFTWARE_APPROVED_DIALOG, { software_name: searchAllTerm});
                await innerDc.beginDialog(SEARCH_SOFTWARE_INSTALLED_DIALOG, { software_name: searchAllTerm});
                await innerDc.beginDialog(SEARCH_SOFTWARE_RAW_DIALOG, { software_name: searchAllTerm});
                await innerDc.beginDialog(SEARCH_SOFTWARE_FINANCIAL_DIALOG, { software_name: searchAllTerm});

                break;

              case 'Log_In':
                // return await innerDc.beginDialog(OAUTH_PROMPT);
                // break;

            }


            // switch (dispatchResults.text) {
            //
            //   case 'ACTO':
            //     await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
            //       //return await innerDc.beginDialog(OAUTH_PROMPT);
            //     break;
            //   case 'FINO':
            //     await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
            //     // return await step.endDialog();
            //     break;
            //   case 'Member':
            //     await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
            //     //return await innerDc.endDialog();
            //     break;
            //   case 'Employer':
            //     await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
            //     //return await innerDc.endDialog();
            //     break;
            //
            // }


        }
    }
}

module.exports.LogoutDialog = LogoutDialog;
