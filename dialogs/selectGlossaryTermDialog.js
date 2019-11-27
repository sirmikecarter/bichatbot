const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet, OAuthPrompt } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer } = require('botbuilder-ai');
const { DialogHelper } = require('./helpers/dialogHelper');
const { SimpleGraphClient } = require('./helpers/simple-graph-client');
const { OAuthHelpers } = require('./helpers/oAuthHelpers');
var arraySort = require('array-sort');
const axios = require('axios');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SelectGlossaryTermDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'selectGlossaryTermDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportNameSearch: [],
          termArray: [],
          userDivision: ''
        };


        this.addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.destinationStep.bind(this)
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
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

          console.log(turnContext.activity.value)

        }
    }

    /**
     * If a destination city has not been provided, prompt for one.
     */
    async destinationStep(stepContext) {

      var tokenResponse = stepContext._info.options.tokenResponse
      var self = this;
      self.state.reportNameSearch = []
      self.state.termArray = []

      console.log(stepContext.context.activity.text)

      switch (stepContext.context.activity.text) {

      case 'Select A Term':

      const client = new SimpleGraphClient(tokenResponse.token);
      const me = await client.getMe();

      await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
            { params: {
              'api-version': '2019-05-06',
              'search': '*'
              },
            headers: {
              'api-key': process.env.GlossarySearchServiceKey,
              'ContentType': 'application/json'
        }

        }).then(response => {

          if (response){


            var itemCount = response.data.value.length
            var itemArray = self.state.reportNameSearch.slice();
            var itemArrayOrg = self.state.reportNameSearch.slice();

            for (var i = 0; i < itemCount; i++)
            {

              const definedBy = response.data.value[i].metadata_definedby.toLowerCase()
              const definedByToken = me.jobTitle.toLowerCase()

              //if (definedBy === definedByToken) {

                const itemResult = response.data.value[i].questions[0]

                if (itemArrayOrg.indexOf(itemResult) === -1)
                {
                  itemArrayOrg.push(itemResult)
                  itemArray.push({'title': itemResult, 'value': itemResult})
                }

              //}

            }
            //console.log(itemArrayOrg)
            self.state.reportNameSearch = arraySort(itemArray, 'title')

         }

        }).catch((error)=>{
               console.log(error);
        });

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createComboListCard('Please Select a Glossary Term', this.state.reportNameSearch, 'glossary_term_selector_value')] });

      break;

      case 'See All Terms':

        const clientNew = new SimpleGraphClient(tokenResponse.token);
        const meNew = await clientNew.getMe();

        const definedByTokenNew = meNew.jobTitle.toLowerCase()

        var self = this;
        self.state.reportNameSearch = []
        self.state.termArray = []

        // Equal to User's Division

        await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
                { params: {
                  'api-version': '2019-05-06',
                  'search': '*',
                  '$filter': 'metadata_definedby eq ' + '\'' + definedByTokenNew + '\''
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

          var card = this.dialogHelper.createGlossaryCard(meNew.jobTitle, data.glossaryterm, data.description, data.definedby, data.output, data.related)

          attachments.push(card);

          }, this)

          await stepContext.context.sendActivity({ attachments: attachments,
          attachmentLayout: AttachmentLayoutTypes.Carousel });

        }else{

          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

        }

        // Not Equal to User's Division

        self.state.termArray = []

        await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
                { params: {
                  'api-version': '2019-05-06',
                  'search': '*',
                  '$filter': 'metadata_definedby ne ' + '\'' + definedByTokenNew + '\''
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

        break;

      }

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

      var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
      await stepContext.context.sendActivity(reply);

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SelectGlossaryTermDialog = SelectGlossaryTermDialog;
