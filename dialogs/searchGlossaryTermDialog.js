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

class SearchGlossaryTermDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'searchGlossaryTermDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportNameSearch: [],
          termArray: [],
          userDivision: ''
        };

        this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
                connectionName: process.env.ConnectionName,
                text: 'Please log in with your credentials and enter the validation code into this chat window to complete the log in process',
                title: 'Log In with My Credentials',
                timeout: 300000
            }))
        this.addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.promptStep.bind(this),
                this.searchStep.bind(this),
                this.searchSpellCheckStep.bind(this),
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

    async promptStep(stepContext) {
        return await stepContext.beginDialog(OAUTH_PROMPT);
    }

    async searchStep(stepContext) {

      const tokenResponse = stepContext.result;

      const clientNew = new SimpleGraphClient(tokenResponse.token);
      const meNew = await clientNew.getMe();

      this.state.userDivision = meNew.jobTitle.toLowerCase()

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...What Should I Search For?','')] });

      return await stepContext.prompt(TEXT_PROMPT, { prompt: '' });
    }

    async searchSpellCheckStep(stepContext) {

      var searchString = stepContext.result;

      const dispatchResults = await this.luisRecognizer.recognize(stepContext.context);

      if (dispatchResults.alteredText)
      {
        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I\'ve noticed a misspelling, click the first box below for updated results or click the second box for your original results','')] });

        return await stepContext.prompt(CHOICE_PROMPT, {
            prompt: '',
            choices: ChoiceFactory.toChoices([dispatchResults.alteredText, searchString])
        });

      } else {
          return await stepContext.next(searchString);
      }

    }

    async searchResultStep(stepContext) {

      var searchString = stepContext.result;

      var self = this;
      self.state.termArray = []

      const dispatchResults = await this.luisRecognizer.recognize(stepContext.context);

      if (dispatchResults.alteredText)
      {
        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I\'ve noticed a misspelling, click the box below for updated results','')] });

        var reply = MessageFactory.suggestedActions([dispatchResults.alteredText]);
        await stepContext.context.sendActivity(reply);
        return await stepContext.endDialog('End Dialog');

      }


      // Equal to User's Division

      await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
              { params: {
                'api-version': '2019-05-06',
                'search': searchString,
                '$filter': 'metadata_definedby eq ' + '\'' + self.state.userDivision + '\''
                },
              headers: {
                'api-key': process.env.GlossarySearchServiceKey,
                'ContentType': 'application/json'
        }

      }).then(response => {

        if (response){

          var itemCount = response.data.value.length

          var itemArray = self.state.termArray.slice();

          for (var i = 0; i < itemCount; i++)
          {
                const glossaryTerm = response.data.value[i].questions[0]
                const glossaryDescription = response.data.value[i].answer
                const glossaryDefinedBy = response.data.value[i].metadata_definedby.toUpperCase()
                const glossaryOutput = response.data.value[i].metadata_output.toUpperCase()
                const glossaryRelated = response.data.value[i].metadata_related

                if (itemArray.indexOf(glossaryTerm) === -1)
                {
                  itemArray.push({'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput, 'related': glossaryRelated})
                }
          }

          self.state.termArray = arraySort(itemArray, 'glossaryterm')


       }

      }).catch((error)=>{
             console.log(error);
      });


      if (this.state.termArray.length > 0){

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.termArray.length + ' Glossary Terms Related to Your Area ','Here are the Results')] });

        var attachments = [];

        this.state.termArray.forEach(function(data){

        var card = this.dialogHelper.createGlossaryCard(data.definedby, data.glossaryterm, data.description, data.definedby, data.output, data.related)

        attachments.push(card);

        }, this)

        await stepContext.context.sendActivity({ attachments: attachments,
        attachmentLayout: AttachmentLayoutTypes.Carousel });

      }else{

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Terms Related to Your Area Were Found','')] });

      }

      // Not Equal to User's Division

      self.state.termArray = []

      await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
              { params: {
                'api-version': '2019-05-06',
                'search': searchString,
                '$filter': 'metadata_definedby ne ' + '\'' + self.state.userDivision + '\''
                },
              headers: {
                'api-key': process.env.GlossarySearchServiceKey,
                'ContentType': 'application/json'
        }

      }).then(response => {

        if (response){

          var itemCount = response.data.value.length

          var itemArray = self.state.termArray.slice();

          for (var i = 0; i < itemCount; i++)
          {
                const glossaryTerm = response.data.value[i].questions[0]
                const glossaryDescription = response.data.value[i].answer
                const glossaryDefinedBy = response.data.value[i].metadata_definedby.toUpperCase()
                const glossaryOutput = response.data.value[i].metadata_output.toUpperCase()
                const glossaryRelated = response.data.value[i].metadata_related

                if (itemArray.indexOf(glossaryTerm) === -1)
                {
                  itemArray.push({'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput, 'related': glossaryRelated})
                }
          }

          self.state.termArray = arraySort(itemArray, 'glossaryterm')


       }

      }).catch((error)=>{
             console.log(error);
      });


      if (this.state.termArray.length > 0){

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.termArray.length + ' Glossary Terms NOT Related to Your Area ','Here are the Results')] });

        var attachments = [];

        this.state.termArray.forEach(function(data){

        var card = this.dialogHelper.createGlossaryCard(data.definedby, data.glossaryterm, data.description, data.definedby, data.output, data.related)

        attachments.push(card);

        }, this)

        await stepContext.context.sendActivity({ attachments: attachments,
        attachmentLayout: AttachmentLayoutTypes.Carousel });

      }else{

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Other Terms Outside of Your Area Were Found','')] });

      }

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

      var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout', 'Search the Glossary']);
      await stepContext.context.sendActivity(reply);

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SearchGlossaryTermDialog = SearchGlossaryTermDialog;
