// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { QnAMaker, LuisRecognizer } = require('botbuilder-ai');
const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet} = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { LuisHelper } = require('./helpers/luisHelper');
const { DialogHelper } = require('./helpers/dialogHelper');
const { SimpleGraphClient } = require('./helpers/simple-graph-client');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SelectReportResultDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'selectReportResultDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportNameSearch: [],
          itemCount: '',
          reportname: '',
          description: '',
          owner: '',
          designee: '',
          approver: '',
          division: '',
          classification: '',
          language: '',
          entities: [],
          keyPhrases: [],
          sentiment: '',
          reportArray: [],
          reportArrayAnalytics: [],
          reportArrayFormData: [],
          reportArrayLanguage: [],
          reportArrayEntities: [],
          reportArrayKeyPhrases: [],
          reportArraySentiment: []
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

      var reportnamequery = "'" + stepContext.context.activity.value.report_name_selector_value + "'"

      var self = this;

      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndex + '/docs?',
              { params: {
                'api-version': '2019-05-06',
                'search': '*',
                '$filter': 'metadata_reportname eq ' + reportnamequery
                },
              headers: {
                'api-key': process.env.SearchServiceKey,
                'ContentType': 'application/json'
        }

      }).then(response => {

        if (response){

         self.state.reportname = response.data.value[0].metadata_reportname
         self.state.description = response.data.value[0].answer
         self.state.owner = response.data.value[0].metadata_owner
         self.state.designee = response.data.value[0].metadata_designee
         self.state.approver = response.data.value[0].metadata_approver
         self.state.division = response.data.value[0].metadata_division
         self.state.classification = response.data.value[0].metadata_classification

       }

      }).catch((error)=>{
             console.log(error);
      });

      //Text Analytics - Body
      var bodyFormData = {
         "documents": [
           {
             "id": "1",
             "text": self.state.description
           }
         ]
       }
      //Text Analytics - Languages
      await axios({
         method: 'post',
         url: 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/languages',
         data: bodyFormData,
         headers: {'Ocp-Apim-Subscription-Key': process.env.TextAnalyticsKey, 'Content-Type': 'application/json'}
         }).then(response => {
             //handle success
             //console.log(response.data.documents[0].detectedLanguages[0].name);
             self.state.language = response.data.documents[0].detectedLanguages[0].name
         }).catch((error)=>{
             //handle error
           console.log(error.response);
       });
       //Text Analytics - Entities
       await axios({
          method: 'post',
          url: 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/entities',
          data: bodyFormData,
          headers: {'Ocp-Apim-Subscription-Key': process.env.TextAnalyticsKey, 'Content-Type': 'application/json'}
          }).then(response => {
              //handle success

              self.state.entities = []
              var itemCount = response.data.documents[0].entities.length

              if(itemCount > 0){
                var itemArray = self.state.entities.slice();

                for (var i = 0; i < itemCount; i++)
                {
                      const itemResult = response.data.documents[0].entities[i].name

                      if (itemArray.indexOf(itemResult) === -1)
                      {
                        itemArray.push(itemResult)
                      }
                }

                self.state.entities = [{'id': 0, 'entities': itemArray }]
                //console.log(self.state.entities)
              }else {
                self.state.entities = [{'id': 0, 'entities': '[No Results]' }]
              }

          }).catch((error)=>{
              //handle error
            console.log(error.response);
        });
        //Text Analytics - Key Phrases
        await axios({
           method: 'post',
           url: 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases',
           data: bodyFormData,
           headers: {'Ocp-Apim-Subscription-Key': process.env.TextAnalyticsKey, 'Content-Type': 'application/json'}
           }).then(response => {
               //handle success

               self.state.keyPhrases = []

               var itemCount = response.data.documents[0].keyPhrases.length

               if (itemCount > 0){
                 var itemArray = self.state.keyPhrases.slice();

                 for (var i = 0; i < itemCount; i++)
                 {
                       const itemResult = response.data.documents[0].keyPhrases[i]

                       if (itemArray.indexOf(itemResult) === -1)
                       {
                         itemArray.push(itemResult)
                       }
                 }

                 self.state.keyPhrases = [{'id': 0, 'keyphrases': itemArray }]
               }else {
                 self.state.keyPhrases = [{'id': 0, 'keyphrases': '[No Results]' }]
               }



           }).catch((error)=>{
               //handle error
             console.log(error.response);
         });
         //Text Analytics - Sentiment
         await axios({
            method: 'post',
            url: 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment',
            data: bodyFormData,
            headers: {'Ocp-Apim-Subscription-Key': process.env.TextAnalyticsKey, 'Content-Type': 'application/json'}
            }).then(response => {
                //handle success
                //console.log(response.data.documents[0].score)
                self.state.sentiment = String(response.data.documents[0].score)


            }).catch((error)=>{
                //handle error
              console.log(error.response);
          });

          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportname, this.state.description, this.state.owner, this.state.designee, this.state.approver, this.state.division, this.state.classification, this.state.language, this.state.entities[0].entities, this.state.keyPhrases[0].keyphrases, this.state.sentiment)] });

          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

          var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
          await stepContext.context.sendActivity(reply);

          return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SelectReportResultDialog = SelectReportResultDialog;
