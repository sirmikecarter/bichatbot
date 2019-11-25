const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./dialogHelper');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SearchSoftwareFinancialDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'searchSoftwareFinancialDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          appArray: [],
          appArrayFinal: [],
          appNotes: [],
          appStatus: []
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


      const searchTerm = stepContext._info.options.software_name;

      var self = this;

      self.state.appArray = []


      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexFinancials + '/docs?',
              { params: {
                'api-version': '2019-05-06',
                'search': searchTerm
                },
              headers: {
                'api-key': process.env.SearchServiceKey,
                'ContentType': 'application/json'
        }

      }).then(response => {

        if (response){

          var itemCount

          if(response.data.value.length === 1){
            itemCount = 1
          }

          if(response.data.value.length === 2){
            itemCount = 2
          }

          if(response.data.value.length === 3){
            itemCount = 3
          }

          if(response.data.value.length > 3){
            itemCount = 3
          }

          var itemArray = self.state.appArray.slice();

          for (var i = 0; i < itemCount; i++)
          {
                const financialScore = i
                const financialId = response.data.value[i].metadata_itemid
                const financialTitle = response.data.value[i].questions[0]
                const financialDesc = response.data.value[i].answer
                const financialYear = response.data.value[i].metadata_year
                const financialContact = response.data.value[i].metadata_contact
                const financialDivision = response.data.value[i].metadata_division
                const financialCost = response.data.value[i].metadata_cost
                const financialApptioCode = response.data.value[i].metadata_apptiocode
                const financialPriorPO = response.data.value[i].metadata_priorpo
                const financialQuantity = response.data.value[i].metadata_quantity

                itemArray.push({'financialScore': financialScore, 'financialId': financialId, 'financialTitle': financialTitle, 'financialDesc': financialDesc, 'financialYear': financialYear, 'financialContact': financialContact, 'financialDivision': financialDivision, 'financialCost': financialCost, 'financialApptioCode': financialApptioCode, 'financialPriorPO': financialPriorPO, 'financialQuantity': financialQuantity})
          }

          self.state.appArray = arraySort(itemArray, 'financialScore')


       }

      }).catch((error)=>{
             console.log(error);
      });

      //console.log(self.state.appArray)

      if (self.state.appArray.length > 0){



        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from the Spending Plan related to ' + searchTerm,'')] });

        var attachments = [];

        this.state.appArray.forEach(function(data){

        var card = this.dialogHelper.createFinancialCard(data.financialId, data.financialTitle, data.financialDesc, data.financialYear, data.financialContact, data.financialDivision, data.financialCost, data.financialApptioCode, data.financialPriorPO, data.financialQuantity)

        attachments.push(card);

        }, this)

        await stepContext.context.sendActivity({ attachments: attachments,
        attachmentLayout: AttachmentLayoutTypes.Carousel });

      }else{

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Items in the Spending Plans Are Related to Your Search','')] });

      }



      return await stepContext.endDialog('End Dialog');
    }

    async resultStep(stepContext) {

      //console.log(stepContext.result.value)

      const teamName = stepContext.result.value.toLowerCase();

      var self = this;

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SearchSoftwareFinancialDialog = SearchSoftwareFinancialDialog;
