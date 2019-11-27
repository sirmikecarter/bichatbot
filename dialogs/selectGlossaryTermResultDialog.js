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

class SelectGlossaryTermResultDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'selectGlossaryTermResultDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          termArray: [],
          glossaryTerm: '',
          glossaryDescription: '',
          glossaryDefinedBy: '',
          glossaryOutput: ''
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

      var glossaryTermQuery = "'" + stepContext.context.activity.value.glossary_term_selector_value + "'"

      const client = new SimpleGraphClient(tokenResponse.token);
      const me = await client.getMe();

      const definedByToken = me.jobTitle.toLowerCase()

      var self = this;
      self.state.termArray = []

      await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
              { params: {
                'api-version': '2019-05-06',
                '$filter': 'search.ismatch(' + glossaryTermQuery + ',' + '\'questions\'' + ')'
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

                  if (stepContext.context.activity.value.glossary_term_selector_value === glossaryTerm ){
                    itemArray.push({'first': 'a', 'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput, 'related': glossaryRelated})
                  }else{
                    itemArray.push({'first': 'z', 'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput, 'related': glossaryRelated})
                  }
                }

          }

          self.state.termArray = arraySort(itemArray, 'first')


       }

      }).catch((error)=>{
             console.log(error);
      });

      if(definedByToken.toUpperCase() === self.state.termArray[0].definedby){
        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...This Term Is Defined By Your Area:','')] });
      }else{
        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...This Term Is NOT Defined By Your Area:','')] });
      }

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(self.state.termArray[0].definedby, self.state.termArray[0].glossaryterm, self.state.termArray[0].description, self.state.termArray[0].definedby, self.state.termArray[0].output, self.state.termArray[0].related)] });



      // console.log(self.state.termArray.length)
      // console.log(self.state.termArray)

      if (self.state.termArray.length > 1){
        console.log(self.state.termArray.length)
        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Also Found a Similar Term:','')] });

        switch (self.state.termArray.length) {

        case 2:

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(self.state.termArray[1].definedby, self.state.termArray[1].glossaryterm, self.state.termArray[1].description, self.state.termArray[1].definedby, self.state.termArray[1].output, self.state.termArray[1].related)] });

        break;


        case 3:

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(self.state.termArray[1].definedby, self.state.termArray[1].glossaryterm, self.state.termArray[1].description, self.state.termArray[1].definedby, self.state.termArray[1].output, self.state.termArray[1].related),
            this.dialogHelper.createGlossaryCard(self.state.termArray[2].definedby, self.state.termArray[2].glossaryterm, self.state.termArray[2].description, self.state.termArray[2].definedby, self.state.termArray[2].output, self.state.termArray[2].related)],
        attachmentLayout: AttachmentLayoutTypes.Carousel });

        break;

        }

      }

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

      var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
      await stepContext.context.sendActivity(reply);

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SelectGlossaryTermResultDialog = SelectGlossaryTermResultDialog;
