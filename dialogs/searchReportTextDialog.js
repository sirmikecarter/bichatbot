const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./dialogHelper');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SearchReportTextDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'searchReportTextDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          reportsSearchString: '*',
          reportsFilterString: '',
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

      const reportApprover = stepContext._info.options.report_approver;
      const reportClassification = stepContext._info.options.report_classification;
      const reportDescription = stepContext._info.options.report_description;
      const reportDesignee = stepContext._info.options.report_designee;
      const reportDivision = stepContext._info.options.report_division;
      const reportName = stepContext._info.options.report_name;
      const reportOwner = stepContext._info.options.report_owner;

      this.state.reportsSearchString = '*'
      this.state.reportsSearchFilterString = ''


      if(reportApprover !== undefined){
        console.log('Report_Approver: ' + reportApprover[0])
        var approverSearch = "'" + String(reportApprover[0]) + "'"

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Approver Filter: ' + approverSearch,'')] });


        if (this.state.reportsSearchFilterString) {
            //console.log('Not Empty')
            this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_approver eq ' + approverSearch
        }else{
            //console.log('Empty')
            this.state.reportsSearchFilterString = 'metadata_approver eq ' + approverSearch
        }

        //await stepContext.context.sendActivity('Report_Approver: ' + dispatchResults.entities.Report_Approver[0])
      }
      if(reportClassification !== undefined){
        console.log('Report_Classification: ' + reportClassification[0])
        var classificationSearch = "'" + String(reportClassification[0]) + "'"

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Classification Filter: ' + classificationSearch,'')] });


        if (this.state.reportsSearchFilterString) {
            //console.log('Not Empty')
            this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_classification eq ' + classificationSearch
        }else{
            //console.log('Empty')
            this.state.reportsSearchFilterString = 'metadata_classification eq ' + classificationSearch
        }

        //await stepContext.context.sendActivity('Report_Classification: ' + dispatchResults.entities.Report_Classification[0])
      }
      if(reportDescription !== undefined){
        console.log('Report_Description: ' + reportDescription[0])

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Searching Reports for: ' + reportDescription[0],'')] });

        this.state.reportsSearchString = String(reportDescription[0])

        //await stepContext.context.sendActivity('Report_Description: ' + dispatchResults.entities.Report_Description[0])
      }
      if(reportDesignee !== undefined){
        console.log('Report_Designee: ' + reportDesignee[0])
        var designeeSearch = "'" + String(reportDesignee[0]) + "'"

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Designee Filter: ' + designeeSearch,'')] });


        if (this.state.reportsSearchFilterString) {
            //console.log('Not Empty')
            this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_designee eq ' + designeeSearch
        }else{
            //console.log('Empty')
            this.state.reportsSearchFilterString = 'metadata_designee eq ' + designeeSearch
        }

        //await stepContext.context.sendActivity('Report_Designee: ' + dispatchResults.entities.Report_Designee[0])
      }
      if(reportDivision !== undefined){
        console.log('Report_Division: ' + reportDivision[0])
        var divisionSearch = "'" + String(reportDivision[0]) + "'"

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Division Filter: ' + divisionSearch,'')] });


        if (this.state.reportsSearchFilterString) {
            //console.log('Not Empty')
            this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_division eq ' + divisionSearch
        }else{
            //console.log('Empty')
            this.state.reportsSearchFilterString = 'metadata_division eq ' + divisionSearch
        }

        //await stepContext.context.sendActivity('Report_Division: ' + dispatchResults.entities.Report_Division[0])
      }
      if(reportName !== undefined){
        console.log('Report_Name: ' + reportName[0])
        var reportNameSearch = "'" + String(reportName[0]) + "'"

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Name Filter: ' + reportNameSearch,'')] });

        if (this.state.reportsSearchFilterString) {
            //console.log('Not Empty')
            this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_reportname eq ' + reportNameSearch
        }else{
            //console.log('Empty')
            this.state.reportsSearchFilterString = 'metadata_reportname eq ' + reportNameSearch
        }

        //await stepContext.context.sendActivity('Report_Name: ' + dispatchResults.entities.Report_Name[0])
      }
      if(reportOwner !== undefined){
        console.log('Report_Owner: ' + reportOwner[0])
        var reportOwnerSearch = "'" + String(reportOwner[0]) + "'"

        await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Owner Filter: ' + reportOwnerSearch,'')] });

        if (this.state.reportsSearchFilterString) {
            //console.log('Not Empty')
            this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_owner eq ' + reportOwnerSearch
        }else{
            //console.log('Empty')
            this.state.reportsSearchFilterString = 'metadata_owner eq ' + reportOwnerSearch
        }

        //await stepContext.context.sendActivity('Report_Owner: ' + dispatchResults.entities.Report_Owner[0])
      }

      console.log('Search String: '+ this.state.reportsSearchString)
      console.log('Filter String: '+ this.state.reportsSearchFilterString)

      if (this.state.reportsSearchFilterString) {
          //console.log('Not Empty')
          var queryParams = {
            'api-version': '2019-05-06',
            'search': this.state.reportsSearchString,
            '$filter': this.state.reportsSearchFilterString
            }
      }else{
          //console.log('Empty')
          var queryParams = {
            'api-version': '2019-05-06',
            'search': this.state.reportsSearchString
            }
      }

      this.state.reportArray = []
      this.state.reportArrayFormData = []
      this.state.reportArrayLanguage = []
      this.state.reportArrayEntities = []
      this.state.reportArrayKeyPhrases = []
      this.state.reportArraySentiment = []

      var self = this;

      await axios.get(process.env.SearchService +'/indexes/'+ process.env.SearchServiceIndex + '/docs?',
        { params: queryParams,
        headers: {
          'api-key': process.env.SearchServiceKey,
          'ContentType': 'application/json'
        }

      }).then(response => {

        if (response){
         // turnContext.sendActivity(`looking through the reports...`);

          var itemCount = response.data.value.length

          var itemArray = self.state.reportArray.slice();

          for (var i = 0; i < itemCount; i++)
          {
                const reportname = response.data.value[i].metadata_reportname
                const description = response.data.value[i].answer
                const owner = response.data.value[i].metadata_owner
                const designee = response.data.value[i].metadata_designee
                const approver = response.data.value[i].metadata_approver
                const division = response.data.value[i].metadata_division
                const classification = response.data.value[i].metadata_classification

                if (itemArray.indexOf(reportname) === -1)
                {
                  itemArray.push({'reportname': reportname, 'description': description, 'owner': owner, 'designee': designee, 'approver': approver, 'division': division, 'classification': classification})
                }
          }

          self.state.reportArray = itemArray

       }

      }).catch((error)=>{
             console.log(error);
      });

      if(this.state.reportArray.length > 0){

      var metaDataOwner = [];
      var metaDataDesignee = [];
      var metaDataApprover = [];
      var metaDataDivision = [];
      var metaDataClassification = [];

      var metaDataOwnerUnique = [];
      var metaDataDesigneeUnique = [];
      var metaDataApproverUnique = [];
      var metaDataDivisionUnique = [];
      var metaDataClassificationUnique = [];

      var metaDataOwnerCount = [];
      var metaDataDesigneeCount = [];
      var metaDataApproverCount = [];
      var metaDataDivisionCount = [];
      var metaDataClassificationCount = [];

      for (var i = 0; i < this.state.reportArray.length; i++)
      {
        metaDataOwner.push(this.state.reportArray[i].owner)
        metaDataDesignee.push(this.state.reportArray[i].designee)
        metaDataApprover.push(this.state.reportArray[i].approver)
        metaDataDivision.push(this.state.reportArray[i].division)
        metaDataClassification.push(this.state.reportArray[i].classification)
      }

      for (var i = 0; i < metaDataOwner.length; i++)
      {

        if (metaDataOwnerUnique.indexOf(metaDataOwner[i]) === -1)
        {
          metaDataOwnerUnique.push(metaDataOwner[i])
        }

        if (metaDataDesigneeUnique.indexOf(metaDataDesignee[i]) === -1)
        {
          metaDataDesigneeUnique.push(metaDataDesignee[i])
        }

        if (metaDataApproverUnique.indexOf(metaDataApprover[i]) === -1)
        {
          metaDataApproverUnique.push(metaDataApprover[i])
        }

        if (metaDataDivisionUnique.indexOf(metaDataDivision[i]) === -1)
        {
          metaDataDivisionUnique.push(metaDataDivision[i])
        }

        if (metaDataClassificationUnique.indexOf(metaDataClassification[i]) === -1)
        {
          metaDataClassificationUnique.push(metaDataClassification[i])
        }

      }
      //
      //
      for (var i = 0; i < metaDataOwnerUnique.length; i++){
        if(metaDataOwnerUnique[i])
        {
          var answerExp = new RegExp(metaDataOwnerUnique[i], 'gi');
          //console.log(metaDataOwner.toString().match(answerExp).length);
          metaDataOwnerCount.push([metaDataOwner.toString().match(answerExp).length, metaDataOwnerUnique[i] ])
        }
      }

      for (var i = 0; i < metaDataDesigneeUnique.length; i++){
        if(metaDataDesigneeUnique[i]){
        var answerExp = new RegExp(metaDataDesigneeUnique[i], 'gi');
        //console.log(metaDataOwner.toString().match(answerExp).length);
        metaDataDesigneeCount.push([metaDataDesignee.toString().match(answerExp).length, metaDataDesigneeUnique[i] ])
        }
      }

      for (var i = 0; i < metaDataApproverUnique.length; i++){
        if(metaDataApproverUnique[i]){
          var answerExp = new RegExp(metaDataApproverUnique[i], 'gi');
          //console.log(metaDataOwner.toString().match(answerExp).length);
          metaDataApproverCount.push([metaDataApprover.toString().match(answerExp).length, metaDataApproverUnique[i] ])
        }
      }

      for (var i = 0; i < metaDataDivisionUnique.length; i++){
        if(metaDataDivisionUnique[i]){
          var answerExp = new RegExp(metaDataDivisionUnique[i], 'gi');
          //console.log(metaDataOwner.toString().match(answerExp).length);
          metaDataDivisionCount.push([metaDataDivision.toString().match(answerExp).length, metaDataDivisionUnique[i] ])
        }
      }

      for (var i = 0; i < metaDataClassificationUnique.length; i++){
        if(metaDataClassificationUnique[i]){
          var answerExp = new RegExp(metaDataClassificationUnique[i], 'gi');
          //console.log(metaDataOwner.toString().match(answerExp).length);
          metaDataClassificationCount.push([metaDataClassification.toString().match(answerExp).length, metaDataClassificationUnique[i] ])
        }
      }


      metaDataOwnerCount = metaDataOwnerCount.sort(this.sortFunction);
      // console.log('Most Seen Owner: ' + metaDataOwnerCount[0][1]);
      // console.log('Most Seen Owner Count: ' + metaDataOwnerCount[0][0]);

      metaDataDesigneeCount = metaDataDesigneeCount.sort(this.sortFunction);
      // console.log('Most Seen Designee: ' + metaDataDesigneeCount[0][1]);
      // console.log('Most Seen Designee Count: ' + metaDataDesigneeCount[0][0]);

      metaDataApproverCount = metaDataApproverCount.sort(this.sortFunction);
      // console.log('Most Seen Approver: ' + metaDataApproverCount[0][1]);
      // console.log('Most Seen Approver Count: ' + metaDataApproverCount[0][0]);

      metaDataDivisionCount = metaDataDivisionCount.sort(this.sortFunction);
      // console.log('Most Seen Division: ' + metaDataDivisionCount[0][1]);
      // console.log('Most Seen Division Count: ' + metaDataDivisionCount[0][0]);

      metaDataClassificationCount = metaDataClassificationCount.sort(this.sortFunction);
      // console.log('Most Seen Classification: ' + metaDataClassificationCount[0][1]);
      // console.log('Most Seen Classification Count: ' + metaDataClassificationCount[0][0]);


      var itemArrayFormData= this.state.reportArrayFormData.slice();

      for (var i = 0; i < this.state.reportArray.length; i++)
      {
        itemArrayFormData.push({'id': i, 'text': this.state.reportArray[i].description})
      }
      this.state.reportArrayFormData = itemArrayFormData


      //Text Analytics - Body
      var bodyFormData = {
         "documents": this.state.reportArrayFormData
       }
      //
       //console.log(JSON.stringify(bodyFormData))

       var self = this;


       //Text Analytics - Languages
       await axios({
          method: 'post',
          url: 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/languages',
          data: bodyFormData,
          headers: {'Ocp-Apim-Subscription-Key': process.env.TextAnalyticsKey, 'Content-Type': 'application/json'}
          }).then(response => {
              //handle success

              var itemCount = self.state.reportArray.length

              if(itemCount > 0){
                var itemArray = self.state.reportArrayLanguage.slice();

                for (var i = 0; i < itemCount; i++)
                {
                      itemArray.push({'id': i, 'language': response.data.documents[i].detectedLanguages[0].name})

                }

                self.state.reportArrayLanguage = itemArray
              }
              //self.state.language = response.data.documents[0].detectedLanguages[0].name
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

               var itemCount = self.state.reportArray.length

               if(itemCount > 0){
                 var itemArray = self.state.reportArrayEntities.slice();

                 for (var i = 0; i < itemCount; i++)
                 {

                   if (response.data.documents[i].entities[0]){
                     //console.log(response.data.documents[i].entities.length)
                     var entitiesArray = []
                     for (var i2 = 0; i2 < response.data.documents[i].entities.length; i2++)
                     {
                      entitiesArray.push(response.data.documents[i].entities[i2].name)
                     }
                     itemArray.push({'id': i, 'entities': entitiesArray})
                   }else {
                     itemArray.push({'id': i, 'entities': 'No Results'})
                   }

                 }

                 self.state.reportArrayEntities = itemArray
                 //console.log(self.state.reportArrayEntities)
               }else {
                 self.state.reportArrayEntities = ['[No Results]']
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

                var itemCount = self.state.reportArray.length

                if (itemCount > 0){
                  var itemArray = self.state.reportArrayKeyPhrases.slice();

                  for (var i = 0; i < itemCount; i++)
                  {
                    if (response.data.documents[i].keyPhrases[0]){

                      var keyphrasesArray = []

                      for (var i2 = 0; i2 < response.data.documents[i].keyPhrases.length; i2++)
                      {
                       keyphrasesArray.push(response.data.documents[i].keyPhrases[i2])
                      }

                      itemArray.push({'id': i, 'keyphrases': keyphrasesArray})
                    }else {
                      itemArray.push({'id': i, 'keyphrases': 'No Results'})
                    }

                  }

                  self.state.reportArrayKeyPhrases = itemArray
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

                 var itemCount = self.state.reportArray.length

                 if(itemCount > 0){
                   var itemArray = self.state.reportArraySentiment.slice();

                   for (var i = 0; i < itemCount; i++)
                   {
                         itemArray.push({'id': i, 'score': String(response.data.documents[i].score)})

                   }

                   self.state.reportArraySentiment = itemArray
                 }


             }).catch((error)=>{
                 //handle error
               console.log(error.response);
           });

         }
            //console.log(this.state.reportArray)
           //  console.log(this.state.reportArrayFormData)
            // console.log(this.state.reportArrayLanguage)
            // console.log(this.state.reportArrayEntities)
            // console.log(this.state.reportArrayKeyPhrases)
            // console.log(this.state.reportArraySentiment)

            // Display Reports
            switch (this.state.reportArray.length) {
            case 0:
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });
                  break;
            case 1:
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score)],
                  attachmentLayout: AttachmentLayoutTypes.Carousel });
                  break;
            case 2:
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                      this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score)],
                  attachmentLayout: AttachmentLayoutTypes.Carousel });
                  break;
            case 3:
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                      this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                      this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score)],
                  attachmentLayout: AttachmentLayoutTypes.Carousel });
                  break;
            case 4:
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                      this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                      this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score),
                      this.dialogHelper.createReportCard(this.state.reportArray[3].reportname, this.state.reportArray[3].description, this.state.reportArray[3].owner, this.state.reportArray[3].designee, this.state.reportArray[3].designee, this.state.reportArray[3].division, this.state.reportArray[3].classification, this.state.reportArrayLanguage[3].language, this.state.reportArrayEntities[3].entities, this.state.reportArrayKeyPhrases[3].keyphrases, this.state.reportArraySentiment[3].score)],
                  attachmentLayout: AttachmentLayoutTypes.Carousel });
                  break;
            case 5:
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                  await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                    this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                    this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score),
                    this.dialogHelper.createReportCard(this.state.reportArray[3].reportname, this.state.reportArray[3].description, this.state.reportArray[3].owner, this.state.reportArray[3].designee, this.state.reportArray[3].designee, this.state.reportArray[3].division, this.state.reportArray[3].classification, this.state.reportArrayLanguage[3].language, this.state.reportArrayEntities[3].entities, this.state.reportArrayKeyPhrases[3].keyphrases, this.state.reportArraySentiment[3].score),
                      this.dialogHelper.createReportCard(this.state.reportArray[4].reportname, this.state.reportArray[4].description, this.state.reportArray[4].owner, this.state.reportArray[4].designee, this.state.reportArray[4].designee, this.state.reportArray[4].division, this.state.reportArray[4].classification, this.state.reportArrayLanguage[4].language, this.state.reportArrayEntities[4].entities, this.state.reportArrayKeyPhrases[4].keyphrases, this.state.reportArraySentiment[4].score)],
                  attachmentLayout: AttachmentLayoutTypes.Carousel });
                break;
            default:
                await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                  this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                  this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score),
                  this.dialogHelper.createReportCard(this.state.reportArray[3].reportname, this.state.reportArray[3].description, this.state.reportArray[3].owner, this.state.reportArray[3].designee, this.state.reportArray[3].designee, this.state.reportArray[3].division, this.state.reportArray[3].classification, this.state.reportArrayLanguage[3].language, this.state.reportArrayEntities[3].entities, this.state.reportArrayKeyPhrases[3].keyphrases, this.state.reportArraySentiment[3].score),
                  this.dialogHelper.createReportCard(this.state.reportArray[4].reportname, this.state.reportArray[4].description, this.state.reportArray[4].owner, this.state.reportArray[4].designee, this.state.reportArray[4].designee, this.state.reportArray[4].division, this.state.reportArray[4].classification, this.state.reportArrayLanguage[4].language, this.state.reportArrayEntities[4].entities, this.state.reportArrayKeyPhrases[4].keyphrases, this.state.reportArraySentiment[4].score)],
                  attachmentLayout: AttachmentLayoutTypes.Carousel });
            }

            if(this.state.reportArray.length > 0){

              await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the top items:','')] });

              await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createMenu('Owner ', metaDataOwnerCount[0][1]),
                this.dialogHelper.createMenu('Designee ', metaDataDesigneeCount[0][1]),
                this.dialogHelper.createMenu('Approver ', metaDataApproverCount[0][1]),
                this.dialogHelper.createMenu('Division ', metaDataDivisionCount[0][1]),
                this.dialogHelper.createMenu('Classification ', metaDataClassificationCount[0][1])],
                attachmentLayout: AttachmentLayoutTypes.Carousel });

                await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

            }


              var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout', 'Search Reports']);
              await stepContext.context.sendActivity(reply);

              return await stepContext.endDialog('End Dialog');

    }

    async resultStep(stepContext) {

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SearchReportTextDialog = SearchReportTextDialog;
