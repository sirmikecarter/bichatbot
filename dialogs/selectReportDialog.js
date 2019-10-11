const { ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');
const { DialogHelper } = require('./dialogHelper');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

const axios = require('axios');

class SelectReportDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'selectReportDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportArray: [],
          reportArrayAnalytics: [],
          reportArrayFormData: [],
          reportArrayLanguage: [],
          reportArrayEntities: [],
          reportArrayKeyPhrases: [],
          reportArraySentiment: [],
          itemArrayMetaUnique: []
        };

        this.addDialog(new TextPrompt(TEXT_PROMPT))
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

      var self = this;
      self.state.reportNameSearch = []

      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndex + '/docs?',
            { params: {
              'api-version': '2019-05-06',
              'search': '*'
              },
            headers: {
              'api-key': process.env.SearchServiceKey,
              'ContentType': 'application/json'
        }

        }).then(response => {

          if (response){

            var itemCount = response.data.value.length
            var itemArray = self.state.reportNameSearch.slice();

            for (var i = 0; i < itemCount; i++)
            {
                  const itemResult = response.data.value[i].metadata_reportname

                  if (itemArray.indexOf(itemResult) === -1)
                  {
                    itemArray.push({'title': itemResult, 'value': itemResult})
                  }
            }

            //console.log(itemArray)
            self.state.reportNameSearch = itemArray

         }

        }).catch((error)=>{
               console.log(error);
        });

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createComboListCard('Please Select a Report', this.state.reportNameSearch, 'report_name_selector_value')] });

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

      var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
      await stepContext.context.sendActivity(reply);

      return await stepContext.endDialog('End Dialog');
    }

}

module.exports.SelectReportDialog = SelectReportDialog;
