const { ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');
const { DialogHelper } = require('./dialogHelper');
const { SimpleGraphClient } = require('../simple-graph-client');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

const axios = require('axios');

class SelectGlossaryTermDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'selectReportDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportNameSearch: [],
          termArray: [],
        };

        this.addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.filterStep.bind(this),
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

    async filterStep(stepContext) {

      return await stepContext.prompt(CHOICE_PROMPT, {
          prompt: 'Single-View or Multi-View?',
          choices: ChoiceFactory.toChoices(['Single-View Glossary', 'Multi-View Glossary'])
      });

    }

    /**
     * If a destination city has not been provided, prompt for one.
     */
    async destinationStep(stepContext, tokenResponse, view) {

      var self = this;
      self.state.reportNameSearch = []
      self.state.termArray = []

      switch (view) {

      case 'Single-View Glossary':

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

            for (var i = 0; i < itemCount; i++)
            {

              const definedBy = response.data.value[i].metadata_definedby.toLowerCase()
              const definedByToken = me.jobTitle.toLowerCase()

              if (definedBy === definedByToken) {

                const itemResult = response.data.value[i].questions[0]

                if (itemArray.indexOf(itemResult) === -1)
                {
                  itemArray.push({'title': itemResult, 'value': itemResult})
                }

              }

            }
            //console.log(itemArray)
            self.state.reportNameSearch = itemArray

         }

        }).catch((error)=>{
               console.log(error);
        });

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createComboListCard('Please Select a Glossary Term', this.state.reportNameSearch, 'glossary_term_selector_value')] });

      break;

      case 'Multi-View Glossary':

        const clientNew = new SimpleGraphClient(tokenResponse.token);
        const meNew = await clientNew.getMe();

        const definedByTokenNew = meNew.jobTitle.toLowerCase()

        var self = this;



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

                  if (itemArray.indexOf(glossaryTerm) === -1)
                  {
                    itemArray.push({'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput})
                  }
            }

            self.state.termArray = arraySort(itemArray, 'glossaryterm')


         }

        }).catch((error)=>{
               console.log(error);
        });

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.termArray.length + ' Glossary Terms ','Here are the Results')] });
        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(meNew.jobTitle, this.state.termArray[0].glossaryterm, this.state.termArray[0].description, this.state.termArray[0].definedby, this.state.termArray[0].output),
            this.dialogHelper.createGlossaryCard(meNew.jobTitle, this.state.termArray[1].glossaryterm, this.state.termArray[1].description, this.state.termArray[1].definedby, this.state.termArray[1].output),
            this.dialogHelper.createGlossaryCard(meNew.jobTitle, this.state.termArray[2].glossaryterm, this.state.termArray[2].description, this.state.termArray[2].definedby, this.state.termArray[2].output),
            this.dialogHelper.createGlossaryCard(meNew.jobTitle, this.state.termArray[3].glossaryterm, this.state.termArray[3].description, this.state.termArray[3].definedby, this.state.termArray[3].output)],
        attachmentLayout: AttachmentLayoutTypes.Carousel });

        break;

        //await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(meNew.jobTitle, this.state.glossaryTerm, this.state.glossaryDescription, this.state.glossaryDefinedBy, this.state.glossaryOutput)] });

      }

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

      var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
      await stepContext.context.sendActivity(reply);

      return await stepContext.endDialog('End Dialog');
    }

}

module.exports.SelectGlossaryTermDialog = SelectGlossaryTermDialog;
