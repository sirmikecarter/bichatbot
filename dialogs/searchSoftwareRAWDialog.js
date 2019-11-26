const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./helpers/dialogHelper');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SearchSoftwareRAWDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'searchSoftwareRAWDialog');

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

      //console.log(searchTerm)

      var self = this;

      self.state.appArray = []

      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexRAW + '/docs?',
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
                const rawScore = i
                const rawIdTitle = response.data.value[i].questions[0]
                const rawName = response.data.value[i].questions[1]
                const rawDesc = response.data.value[i].answer
                const rawCategory = response.data.value[i].metadata_requestcategory
                const rawCategoryOther = response.data.value[i].metadata_requestcategoryother
                const rawPhase = response.data.value[i].metadata_requestphase
                const rawType = response.data.value[i].metadata_requesttype
                const rawBizLine = response.data.value[i].metadata_businessline
                const rawSubmitter = response.data.value[i].metadata_submittername
                const rawSubmitterDiv = response.data.value[i].metadata_submitterdivision
                const rawSubmitterUnit = response.data.value[i].metadata_submitterunit
                const rawOwner = response.data.value[i].metadata_owner
                const rawOwnerDiv = response.data.value[i].metadata_ownerdivision
                const rawOwnerUnit = response.data.value[i].metadata_ownerunit
                const rawDateSubmit = response.data.value[i].metadata_datesubmitted
                const rawDateComplete = response.data.value[i].metadata_datecompleted
                const rawId = response.data.value[i].metadata_rawid

                itemArray.push({'rawScore': rawScore, 'rawIdTitle': rawIdTitle, 'rawName': rawName, 'rawDesc': rawDesc, 'rawCategory': rawCategory, 'rawCategoryOther': rawCategoryOther, 'rawPhase': rawPhase, 'rawType': rawType, 'rawBizLine': rawBizLine, 'rawSubmitter': rawSubmitter, 'rawSubmitterDiv': rawSubmitterDiv, 'rawSubmitterUnit': rawSubmitterUnit, 'rawOwner': rawOwner, 'rawOwnerDiv': rawOwnerDiv, 'rawOwnerUnit': rawOwnerUnit, 'rawDateSubmit': rawDateSubmit, 'rawDateComplete': rawDateComplete, 'rawId': rawId})
          }

          self.state.appArray = arraySort(itemArray, 'rawScore')


       }

      }).catch((error)=>{
             console.log(error);
      });

      //console.log(self.state.appArray)

      if (self.state.appArray.length > 0){



        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from the RAW system related to ' + searchTerm,'')] });

        var attachments = [];

        this.state.appArray.forEach(function(data){

        var card = this.dialogHelper.createRAWCard(data.rawIdTitle, data.rawName, data.rawDesc, data.rawCategory, data.rawCategoryOther, data.rawPhase, data.rawType, data.rawBizLine, data.rawSubmitter, data.rawSubmitterDiv, data.rawSubmitterUnit, data.rawOwner, data.rawOwnerDiv, data.rawOwnerUnit, data.rawDateSubmit, data.rawDateComplete, data.rawId)

        attachments.push(card);

        }, this)

        await stepContext.context.sendActivity({ attachments: attachments,
        attachmentLayout: AttachmentLayoutTypes.Carousel });

      }else{

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No RAWs Related to Your Search Were Found','')] });

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

module.exports.SearchSoftwareRAWDialog = SearchSoftwareRAWDialog;
