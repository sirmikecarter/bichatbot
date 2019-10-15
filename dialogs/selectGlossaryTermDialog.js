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

class SelectGlossaryTermDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'selectGlossaryTermDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportNameSearch: [],
          termArray: [],
          userDivision: ''
        };

    }

    /**
     * If a destination city has not been provided, prompt for one.
     */
    async destinationStep(stepContext, tokenResponse, view) {

      var self = this;
      self.state.reportNameSearch = []
      self.state.termArray = []

      switch (view) {

      case 'Select 1 Term':

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
