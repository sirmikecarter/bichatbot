const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./dialogHelper');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SearchSoftwareApprovedDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'searchSoftwareApprovedDialog');

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

      var self = this;

      const searchTerm = stepContext._info.options.software_name;

      //console.log('SOFTWARE NAME: ' + stepContext._info.options.software_name)

      var self = this;

      self.state.appArray = []
      self.state.appNotes = []
      self.state.appArrayFinal = []
      self.state.appStatus = []

      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexApproved + '/docs?',
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
                const appScore = i
                const appName = response.data.value[i].questions[0]
                const appDesc = response.data.value[i].answer
                const appType = response.data.value[i].metadata_type
                const appId = response.data.value[i].metadata_provisionid

                itemArray.push({'appScore': appScore, 'appName': appName, 'appDesc': appDesc, 'appType': appType, 'appId': appId})
          }

          self.state.appArray = arraySort(itemArray, 'appScore')


       }

      }).catch((error)=>{
             console.log(error);
      });

      //console.log(this.state.appArray)


      var itemArrayFinal = self.state.appArrayFinal.slice();

      for (var i = 0; i < self.state.appArray.length; i++)
      {

        self.state.appNotes = []
        self.state.appStatus = []


      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexApprovedStatus + '/docs?',
              { params: {
                'api-version': '2019-05-06',
                'search': '*',
                '$filter': 'metadata_provisionid eq ' + '\'' + self.state.appArray[i].appId + '\''
                },
              headers: {
                'api-key': process.env.SearchServiceKey,
                'ContentType': 'application/json'
        }

      }).then(response => {

        if (response){

        //console.log(response.data.value[0].@search.score)


          var noteCount = response.data.value.length

          var noteArray = self.state.appNotes.slice();
          var statusArray = self.state.appStatus.slice();

          for (var i2 = 0; i2 < noteCount; i2++)
          {
                const appNotes = response.data.value[i2].answer
                const appStatus = response.data.value[i2].questions[0]
                const appStatusDate = response.data.value[i2].metadata_statusdate
                const appStatusValue = response.data.value[i2].metadata_statusvalue

                if (noteArray.indexOf(appNotes) === -1 && appNotes !== 'undefined')
                {
                noteArray.push(appNotes)
                }



                if (appStatusValue === '1')
                {
                  statusArray.push({'appStatus': appStatus, 'appStatusDate': appStatusDate, 'appStatusValue': appStatusValue})
                }




          }

          self.state.appNotes = noteArray
          self.state.appStatus = statusArray

          //console.log(statusArray)

          itemArrayFinal.push({'appScore': self.state.appArray[i].appScore, 'appName': self.state.appArray[i].appName, 'appDesc': self.state.appArray[i].appDesc, 'appType': self.state.appArray[i].appType, 'appId': self.state.appArray[i].appId, 'appStatus': self.state.appStatus[0].appStatus,'appStatusDate': self.state.appStatus[0].appStatusDate, 'appNote1': self.state.appNotes[0], 'appNote2': self.state.appNotes[1], 'appNote3': self.state.appNotes[2]})



       }

      }).catch((error)=>{
             console.log(error);
      });


    }

    self.state.appArrayFinal = arraySort(itemArrayFinal, 'appScore')



    //console.log(self.state.appArrayFinal)


      if (self.state.appArrayFinal.length > 0){


        var answerExp1 = self.state.appArrayFinal[0].appName.toLowerCase().replace("[", "");
        var answerExp2 = answerExp1.toLowerCase().replace("]", "");

        var approveCheck = answerExp2.toLowerCase().includes(String(searchTerm));

        //console.log(approveCheck)

        if (approveCheck === false && self.state.appArrayFinal[1]){
          answerExp1 = self.state.appArrayFinal[1].appName.toLowerCase().replace("[", "");
          answerExp2 = answerExp1.toLowerCase().replace("]", "");
          approveCheck = answerExp2.toLowerCase().includes(String(searchTerm));
        }

        //console.log(approveCheck)

        if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Current')
        {
          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchTerm + ' is Approved to Use ','')] });
        }

        if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Restricted')
        {
          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchTerm + ' is Approved to Use but is Restricted. Check the Notes tab for the Restriction Note ','')] });
        }

        if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Experimental')
        {
          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchTerm + ' is Approved to Use but for Experimental Purposes Only ','')] });
        }

        if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Retired')
        {
          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No, it appears ' + searchTerm + ' is Retired and No longer approved to Use ','')] });
        }

        if (approveCheck === true && self.state.appArrayFinal[0].appStatus === 'Sunset')
        {
          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Yes, it appears ' + searchTerm + ' is Approved but will soon reach end of life ','')] });
        }


        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from our Application Portfolio related to ' + searchTerm,'')] });

        var attachments = [];

        this.state.appArrayFinal.forEach(function(data){

        var card = this.dialogHelper.createAppApprovalCard(data.appName, data.appDesc, data.appType, data.appId, data.appStatus, data.appStatusDate, data.appNote1, data.appNote2, data.appNote3)

        attachments.push(card);

        }, this)

        await stepContext.context.sendActivity({ attachments: attachments,
        attachmentLayout: AttachmentLayoutTypes.Carousel });

      }else{

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Software Applications Related to Your Search Were Found','')] });

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

module.exports.SearchSoftwareApprovedDialog = SearchSoftwareApprovedDialog;
