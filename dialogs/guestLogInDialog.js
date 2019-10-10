const { ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet, OAuthPrompt } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer } = require('botbuilder-ai');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');
const { DialogHelper } = require('./dialogHelper');
const { SimpleGraphClient } = require('../simple-graph-client');
const { OAuthHelpers } = require('../oAuthHelpers');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';

const axios = require('axios');

class GuestLogInDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'guestLogInDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportNameSearch: [],
          termArray: [],
          userDivision: '',
          itemCount: ''
        };

        this.addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.typeOfUserStep.bind(this),
                this.userDialog.bind(this),
                this.searchResultStep.bind(this)
            ]));

        this.initialDialogId = WATERFALL_DIALOG;

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

        this.luisRecognizer = new LuisRecognizer(luisApplication, luisPredictionOptions);

    }

    /**
     * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} turnContext
     * @param {*} accessor
     */
    async onTurn(turnContext, accessor) {
        // Call QnA Maker and get results.
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        await dialogContext.beginDialog(this.id);

        if (turnContext.activity.value){

          //console.log(turnContext.activity.value)

        }
    }

    async typeOfUserStep(stepContext) {

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...What Type of User Are You?','')] });

      return await stepContext.prompt(CHOICE_PROMPT, {
          prompt: '',
          choices: ChoiceFactory.toChoices(['CalPERS Staff', 'Member', 'Employer'])
      });
    }

    async userDialog(stepContext) {

      var userString = stepContext.result.value;

      console.log(userString)

      var self = this

      switch (userString) {

      case 'CalPERS Staff':

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Dont Forget about the Chilli Cook-off and Halloween Costume Contest','Its on Thursday, October 31st from 11am to 1pm, in the LPN 1st Floor Atrium')] });
      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createImageCard()] });
      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Im dressing up as a BOT for the Costume Contest !','')] });

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...What Division Are You From?','')] });

      return await stepContext.prompt(CHOICE_PROMPT, {
          prompt: '',
          choices: ChoiceFactory.toChoices(['ACTO', 'FINO'])
      });

      break;

      case 'Member':

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...CalHR now has a a Benefits Calculator advailable for you to see what your costs will be for health, dental and vision benefits based on your plan choices ','')] });
      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createLink('Click to this link for more information','https://eservices.calhr.ca.gov/BenefitsCalculatorExternal/')] });


      return await stepContext.endDialog('End Dialog');

      break;

      case 'Employer':

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...The deadline for processing all Open Enrollment transactions is 11:59 p.m. (Pacific Time) on Friday, November 1, 2019. Changes made during Open Enrollment take effect January 1, 2020.','')] });
      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createLink('View Employer Resources','https://www.calpers.ca.gov/page/employers/benefit-programs/health-benefits/open-enrollment-for-employers')] });
      return await stepContext.endDialog('End Dialog');
      break;


      default:

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Didnt Understand That','')] });
      return await stepContext.endDialog('End Dialog');


      break;




    }


    }

    async searchResultStep(stepContext) {

      var searchString = stepContext.result.value;

      //console.log(searchString)

      const definedByToken = searchString.toLowerCase()
      var self = this;

      await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
              { params: {
                'api-version': '2019-05-06',
                'search': '*',
                '$filter': 'metadata_definedby eq ' + '\'' + definedByToken + '\''
                },
              headers: {
                'api-key': process.env.GlossarySearchServiceKey,
                'ContentType': 'application/json'
        }

      }).then(response => {

        if (response){

          self.state.itemCount = response.data.value.length


       }

      }).catch((error)=>{
             console.log(error);
      });


      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I see there are ' + self.state.itemCount + ' terms in the business glossary that are defined by your division','Please Login to see additional information')] });

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.GuestLogInDialog = GuestLogInDialog;
