const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./helpers/dialogHelper');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SearchSoftwareInstalledDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'searchSoftwareInstalledDialog');

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

      var self = this;

      self.state.appArray = []

      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndexInstalled + '/docs?',
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
                const appClass = response.data.value[i].metadata_classification
                const appPublisher = response.data.value[i].metadata_publisher
                const appVersion = response.data.value[i].metadata_version
                const appEdition = response.data.value[i].metadata_edition
                const appCategory = response.data.value[i].metadata_softwarecategory
                const appSubCategory = response.data.value[i].metadata_softwaresubcategory
                const appInstalled = response.data.value[i].metadata_installed
                const appReleaseDate = response.data.value[i].metadata_releasedate
                const appEndOfSales = response.data.value[i].metadata_endofsales
                const appEndofLife = response.data.value[i].metadata_endoflife
                const appEndOfSupport = response.data.value[i].metadata_endofsupport
                const appEndofExtendedSupport = response.data.value[i].metadata_endofextendedsupport
                const appId = response.data.value[i].metadata_flexeraid

                itemArray.push({'appScore': appScore, 'appName': appName, 'appClass': appClass, 'appPublisher': appPublisher, 'appVersion': appVersion, 'appEdition': appEdition, 'appCategory': appCategory, 'appSubCategory': appSubCategory, 'appInstalled': appInstalled, 'appReleaseDate': appReleaseDate, 'appEndOfSales': appEndOfSales, 'appEndofLife': appEndofLife, 'appEndOfSupport': appEndOfSupport, 'appEndofExtendedSupport': appEndofExtendedSupport, 'appId': appId})
          }

          self.state.appArray = arraySort(itemArray, 'appScore')


       }

      }).catch((error)=>{
             console.log(error);
      });

      //console.log(self.state.appArray)


      if (self.state.appArray.length > 0){



        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the Top Results from Flexera related to ' + searchTerm,'')] });

        var attachments = [];

        this.state.appArray.forEach(function(data){

        var card = this.dialogHelper.createAppInstalledCard(data.appName, data.appClass, data.appId, data.appInstalled, data.appCategory, data.appSubCategory, data.appStatusDate, data.appPublisher, data.appVersion, data.appEdition, data.appReleaseDate, data.appEndOfSales, data.appEndofLife, data.appEndOfSupport, data.appEndofExtendedSupport)

        attachments.push(card);

        }, this)

        await stepContext.context.sendActivity({ attachments: attachments,
        attachmentLayout: AttachmentLayoutTypes.Carousel });

      }else{

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Installed Applications Related to Your Search Were Found','')] });

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

module.exports.SearchSoftwareInstalledDialog = SearchSoftwareInstalledDialog;
