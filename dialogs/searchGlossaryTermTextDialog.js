const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./helpers/dialogHelper');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SearchGlossaryTermTextDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'searchGlossaryTermTextDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          termArray: []
        };


        this.addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new ChoicePrompt(CHOICE_PROMPT))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.destinationStep.bind(this),
                this.resultStep.bind(this)
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

      var self = this;

      const searchTerm = stepContext._info.options.glossary_term;

      self.state.termArray = []
      //return await stepContext.beginDialog(SEARCH_GLOSSARY_TERM_DIALOG);

      if(searchTerm !== undefined){
        console.log('Term: ' + searchTerm)
        var termSearch = "'" + String(searchTerm) + "'"

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Searching Business Glossary for: ' + searchTerm,'')] });

        await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
                { params: {
                  'api-version': '2019-05-06',
                  'search': termSearch
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

          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.termArray.length + ' Glossary Terms ','Here are the Results')] });

          var attachments = [];

          this.state.termArray.forEach(function(data){

          var card = this.dialogHelper.createGlossaryCard(data.definedby, data.glossaryterm, data.description, data.definedby, data.output, data.related)

          attachments.push(card);

          }, this)

          await stepContext.context.sendActivity({ attachments: attachments,
          attachmentLayout: AttachmentLayoutTypes.Carousel });

        }else{

          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

        }

      }else{

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

      }

      return await stepContext.endDialog('End Dialog');
    }

    async resultStep(stepContext) {

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SearchGlossaryTermTextDialog = SearchGlossaryTermTextDialog;
