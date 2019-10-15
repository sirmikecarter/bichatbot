// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { ComponentDialog, ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer, QnAMaker } = require('botbuilder-ai');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { ActivityTypes } = require('botbuilder');
const { DialogHelper } = require('./dialogHelper');
const { GuestLogInDialog } = require('./guestLogInDialog');
const { SearchGlossaryTermDialog } = require('./searchGlossaryTermDialog');
const { SearchReportDialog } = require('./searchReportDialog');
const axios = require('axios');
var arraySort = require('array-sort');


const CHOICE_PROMPT = 'choicePrompt';
const OAUTH_PROMPT = 'oAuthPrompt';
const GUEST_LOG_IN_DIALOG = 'guestLogInDialog';
const SEARCH_GLOSSARY_TERM_DIALOG = 'searchGlossaryTermDialog';
const SEARCH_REPORT_DIALOG = 'searchReportDialog';
const QNA_TOP_N = 1;
const QNA_CONFIDENCE_THRESHOLD = 0.5;


class LogoutDialog extends ComponentDialog {

  constructor(id) {
      super(id || 'logoutDialog');

      this.state = {
        cityName: '',
        cityTemp: '',
        cityTempHi: '',
        cityTempLo: '',
        teamId: '',
        teamName: '',
        teamBadge: '',
        homeScore: '',
        homeTeam: '',
        homeTeamId: '',
        homeTeamBadge: '',
        awayScore: '',
        awayTeam: '',
        awayTeamId: '',
        awayTeamBadge: '',
        dateEvent: '',
        reportsSearchString: '*',
        reportsFilterString: '',
        reportArray: [],
        reportArrayAnalytics: [],
        reportArrayFormData: [],
        reportArrayLanguage: [],
        reportArrayEntities: [],
        reportArrayKeyPhrases: [],
        reportArraySentiment: [],
        itemArrayMetaUnique: [],
        termArray: []
      };

      const luisApplication = {
          applicationId: process.env.LuisAppId,
          azureRegion: process.env.LuisAPIHostName,
          // CAUTION: Authoring key is used in this example as it is appropriate for prototyping.
          // When implimenting for deployment/production, assign and use a subscription key instead of an authoring key.
          endpointKey: process.env.LuisAPIKey
      };

      const luisPredictionOptions = {
          spellCheck: true,
          bingSpellCheckSubscriptionKey: process.env.BingSpellCheck

      };

      this.qnaRecognizer = new QnAMaker({
          knowledgeBaseId: process.env.QnAKbId,
          endpointKey: process.env.QnAEndpointKey,
          host: process.env.QnAHostname
      });

      this.luisRecognizer = new LuisRecognizer(luisApplication, luisPredictionOptions);

  }

    async onBeginDialog(innerDc, options) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onBeginDialog(innerDc, options);
    }

    async onContinueDialog(innerDc) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onContinueDialog(innerDc);
    }

    async interrupt(innerDc) {
        if (innerDc.context.activity.type === ActivityTypes.Message) {
            const text = innerDc.context.activity.text ? innerDc.context.activity.text.toLowerCase() : '';
            if (text === 'logout') {
                // The bot adapter encapsulates the authentication processes.
                const botAdapter = innerDc.context.adapter;
                await botAdapter.signOutUser(innerDc.context, process.env.ConnectionName);

                await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('You have been signed out.','')] });

                await innerDc.prompt(CHOICE_PROMPT, {
                    prompt: '',
                    choices: ChoiceFactory.toChoices(['Log In'])
                });

                //await innerDc.context.sendActivity('You have been signed out.');
                return await innerDc.cancelAllDialogs();
            }

            if (text === 'menu') {
              console.log(text)

            }

            const dispatchResults = await this.luisRecognizer.recognize(innerDc.context);
            const dispatchTopIntent = LuisRecognizer.topIntent(dispatchResults);

            console.log(dispatchTopIntent)

            console.log(dispatchResults.text)


            switch (dispatchTopIntent) {
              case 'General':
                //console.log(dispatchResults.intents)
                //console.log(dispatchTopIntent)
                const qnaResult = await this.qnaRecognizer.generateAnswer(dispatchResults.text, QNA_TOP_N, QNA_CONFIDENCE_THRESHOLD);
                if (!qnaResult || qnaResult.length === 0 || !qnaResult[0].answer){
                  //await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                }else{
                  await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard(String(qnaResult[0].answer),'')] });
                }
                //console.log(qnaResult[0].answer)
                break;
              case 'Glossary':
                // console.log(dispatchResults)
                // console.log(dispatchResults.entities.Term[0])
                var self = this;
                self.state.termArray = []
                //return await innerDc.beginDialog(SEARCH_GLOSSARY_TERM_DIALOG);

                if(dispatchResults.entities.Term !== undefined){
                  console.log('Term: ' + dispatchResults.entities.Term[0])
                  var termSearch = "'" + String(dispatchResults.entities.Term[0]) + "'"

                  await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Searching Business Glossary for: ' + dispatchResults.entities.Term[0],'')] });

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

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.termArray.length + ' Glossary Terms ','Here are the Results')] });

                    var attachments = [];

                    this.state.termArray.forEach(function(data){

                    var card = this.dialogHelper.createGlossaryCard(data.definedby, data.glossaryterm, data.description, data.definedby, data.output, data.related)

                    attachments.push(card);

                    }, this)

                    await innerDc.context.sendActivity({ attachments: attachments,
                    attachmentLayout: AttachmentLayoutTypes.Carousel });

                  }else{

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

                  }



                  //await innerDc.context.sendActivity('Report_Approver: ' + dispatchResults.entities.Report_Approver[0])
                }else{

                  await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

                }



                break;
              case 'Reports':
                  //console.log(dispatchResults)

                  this.state.reportsSearchString = '*'
                  this.state.reportsSearchFilterString = ''


                  if(dispatchResults.entities.Report_Approver !== undefined){
                    console.log('Report_Approver: ' + dispatchResults.entities.Report_Approver[0])
                    var approverSearch = "'" + String(dispatchResults.entities.Report_Approver[0]) + "'"

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Approver Filter: ' + dispatchResults.entities.Report_Approver[0],'')] });


                    if (this.state.reportsSearchFilterString) {
                        //console.log('Not Empty')
                        this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_approver eq ' + approverSearch
                    }else{
                        //console.log('Empty')
                        this.state.reportsSearchFilterString = 'metadata_approver eq ' + approverSearch
                    }

                    //await innerDc.context.sendActivity('Report_Approver: ' + dispatchResults.entities.Report_Approver[0])
                  }
                  if(dispatchResults.entities.Report_Classification !== undefined){
                    console.log('Report_Classification: ' + dispatchResults.entities.Report_Classification[0])
                    var classificationSearch = "'" + String(dispatchResults.entities.Report_Classification[0]) + "'"

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Classification Filter: ' + dispatchResults.entities.Report_Classification[0],'')] });


                    if (this.state.reportsSearchFilterString) {
                        //console.log('Not Empty')
                        this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_classification eq ' + classificationSearch
                    }else{
                        //console.log('Empty')
                        this.state.reportsSearchFilterString = 'metadata_classification eq ' + classificationSearch
                    }

                    //await innerDc.context.sendActivity('Report_Classification: ' + dispatchResults.entities.Report_Classification[0])
                  }
                  if(dispatchResults.entities.Report_Description !== undefined){
                    console.log('Report_Description: ' + dispatchResults.entities.Report_Description[0])

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Searching Reports for: ' + dispatchResults.entities.Report_Description[0],'')] });

                    this.state.reportsSearchString = String(dispatchResults.entities.Report_Description[0])

                    //await innerDc.context.sendActivity('Report_Description: ' + dispatchResults.entities.Report_Description[0])
                  }
                  if(dispatchResults.entities.Report_Designee !== undefined){
                    console.log('Report_Designee: ' + dispatchResults.entities.Report_Designee[0])
                    var designeeSearch = "'" + String(dispatchResults.entities.Report_Designee[0]) + "'"

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Designee Filter: ' + dispatchResults.entities.Report_Designee[0],'')] });


                    if (this.state.reportsSearchFilterString) {
                        //console.log('Not Empty')
                        this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_designee eq ' + designeeSearch
                    }else{
                        //console.log('Empty')
                        this.state.reportsSearchFilterString = 'metadata_designee eq ' + designeeSearch
                    }

                    //await innerDc.context.sendActivity('Report_Designee: ' + dispatchResults.entities.Report_Designee[0])
                  }
                  if(dispatchResults.entities.Report_Division !== undefined){
                    console.log('Report_Division: ' + dispatchResults.entities.Report_Division[0])
                    var divisionSearch = "'" + String(dispatchResults.entities.Report_Division[0]) + "'"

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Division Filter: ' + dispatchResults.entities.Report_Division[0],'')] });


                    if (this.state.reportsSearchFilterString) {
                        //console.log('Not Empty')
                        this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_division eq ' + divisionSearch
                    }else{
                        //console.log('Empty')
                        this.state.reportsSearchFilterString = 'metadata_division eq ' + divisionSearch
                    }

                    //await innerDc.context.sendActivity('Report_Division: ' + dispatchResults.entities.Report_Division[0])
                  }
                  if(dispatchResults.entities.Report_Name !== undefined){
                    console.log('Report_Name: ' + dispatchResults.entities.Report_Name[0])
                    var reportNameSearch = "'" + String(dispatchResults.entities.Report_Name[0]) + "'"

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Name Filter: ' + dispatchResults.entities.Report_Name[0],'')] });

                    if (this.state.reportsSearchFilterString) {
                        //console.log('Not Empty')
                        this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_reportname eq ' + reportNameSearch
                    }else{
                        //console.log('Empty')
                        this.state.reportsSearchFilterString = 'metadata_reportname eq ' + reportNameSearch
                    }

                    //await innerDc.context.sendActivity('Report_Name: ' + dispatchResults.entities.Report_Name[0])
                  }
                  if(dispatchResults.entities.Report_Owner !== undefined){
                    console.log('Report_Owner: ' + dispatchResults.entities.Report_Owner[0])
                    var reportOwner = "'" + String(dispatchResults.entities.Report_Owner[0]) + "'"

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('Report Owner Filter: ' + dispatchResults.entities.Report_Owner[0],'')] });

                    if (this.state.reportsSearchFilterString) {
                        //console.log('Not Empty')
                        this.state.reportsSearchFilterString = this.state.reportsSearchFilterString + ' and ' + 'metadata_owner eq ' + reportOwner
                    }else{
                        //console.log('Empty')
                        this.state.reportsSearchFilterString = 'metadata_owner eq ' + reportOwner
                    }

                    //await innerDc.context.sendActivity('Report_Owner: ' + dispatchResults.entities.Report_Owner[0])
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
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });
                              break;
                        case 1:
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score)],
                              attachmentLayout: AttachmentLayoutTypes.Carousel });
                              break;
                        case 2:
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                                  this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score)],
                              attachmentLayout: AttachmentLayoutTypes.Carousel });
                              break;
                        case 3:
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                                  this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                                  this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score)],
                              attachmentLayout: AttachmentLayoutTypes.Carousel });
                              break;
                        case 4:
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                                  this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                                  this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score),
                                  this.dialogHelper.createReportCard(this.state.reportArray[3].reportname, this.state.reportArray[3].description, this.state.reportArray[3].owner, this.state.reportArray[3].designee, this.state.reportArray[3].designee, this.state.reportArray[3].division, this.state.reportArray[3].classification, this.state.reportArrayLanguage[3].language, this.state.reportArrayEntities[3].entities, this.state.reportArrayKeyPhrases[3].keyphrases, this.state.reportArraySentiment[3].score)],
                              attachmentLayout: AttachmentLayoutTypes.Carousel });
                              break;
                        case 5:
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                              await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                                this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                                this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score),
                                this.dialogHelper.createReportCard(this.state.reportArray[3].reportname, this.state.reportArray[3].description, this.state.reportArray[3].owner, this.state.reportArray[3].designee, this.state.reportArray[3].designee, this.state.reportArray[3].division, this.state.reportArray[3].classification, this.state.reportArrayLanguage[3].language, this.state.reportArrayEntities[3].entities, this.state.reportArrayKeyPhrases[3].keyphrases, this.state.reportArraySentiment[3].score),
                                  this.dialogHelper.createReportCard(this.state.reportArray[4].reportname, this.state.reportArray[4].description, this.state.reportArray[4].owner, this.state.reportArray[4].designee, this.state.reportArray[4].designee, this.state.reportArray[4].division, this.state.reportArray[4].classification, this.state.reportArrayLanguage[4].language, this.state.reportArrayEntities[4].entities, this.state.reportArrayKeyPhrases[4].keyphrases, this.state.reportArraySentiment[4].score)],
                              attachmentLayout: AttachmentLayoutTypes.Carousel });
                            break;
                        default:
                            await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...I Found ' + this.state.reportArray.length + ' Reports ','Here are the Top Results')] });
                            await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createReportCard(this.state.reportArray[0].reportname, this.state.reportArray[0].description, this.state.reportArray[0].owner, this.state.reportArray[0].designee, this.state.reportArray[0].designee, this.state.reportArray[0].division, this.state.reportArray[0].classification, this.state.reportArrayLanguage[0].language, this.state.reportArrayEntities[0].entities, this.state.reportArrayKeyPhrases[0].keyphrases, this.state.reportArraySentiment[0].score),
                              this.dialogHelper.createReportCard(this.state.reportArray[1].reportname, this.state.reportArray[1].description, this.state.reportArray[1].owner, this.state.reportArray[1].designee, this.state.reportArray[1].designee, this.state.reportArray[1].division, this.state.reportArray[1].classification, this.state.reportArrayLanguage[1].language, this.state.reportArrayEntities[1].entities, this.state.reportArrayKeyPhrases[1].keyphrases, this.state.reportArraySentiment[1].score),
                              this.dialogHelper.createReportCard(this.state.reportArray[2].reportname, this.state.reportArray[2].description, this.state.reportArray[2].owner, this.state.reportArray[2].designee, this.state.reportArray[2].designee, this.state.reportArray[2].division, this.state.reportArray[2].classification, this.state.reportArrayLanguage[2].language, this.state.reportArrayEntities[2].entities, this.state.reportArrayKeyPhrases[2].keyphrases, this.state.reportArraySentiment[2].score),
                              this.dialogHelper.createReportCard(this.state.reportArray[3].reportname, this.state.reportArray[3].description, this.state.reportArray[3].owner, this.state.reportArray[3].designee, this.state.reportArray[3].designee, this.state.reportArray[3].division, this.state.reportArray[3].classification, this.state.reportArrayLanguage[3].language, this.state.reportArrayEntities[3].entities, this.state.reportArrayKeyPhrases[3].keyphrases, this.state.reportArraySentiment[3].score),
                              this.dialogHelper.createReportCard(this.state.reportArray[4].reportname, this.state.reportArray[4].description, this.state.reportArray[4].owner, this.state.reportArray[4].designee, this.state.reportArray[4].designee, this.state.reportArray[4].division, this.state.reportArray[4].classification, this.state.reportArrayLanguage[4].language, this.state.reportArrayEntities[4].entities, this.state.reportArrayKeyPhrases[4].keyphrases, this.state.reportArraySentiment[4].score)],
                              attachmentLayout: AttachmentLayoutTypes.Carousel });
                        }

                        if(this.state.reportArray.length > 0){

                          await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Here are the top items:','')] });

                          await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createMenu('Owner ', metaDataOwnerCount[0][1]),
                            this.dialogHelper.createMenu('Designee ', metaDataDesigneeCount[0][1]),
                            this.dialogHelper.createMenu('Approver ', metaDataApproverCount[0][1]),
                            this.dialogHelper.createMenu('Division ', metaDataDivisionCount[0][1]),
                            this.dialogHelper.createMenu('Classification ', metaDataClassificationCount[0][1])],
                            attachmentLayout: AttachmentLayoutTypes.Carousel });

                            await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

                        }


                          var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout', 'Search Reports']);
                          await innerDc.context.sendActivity(reply);

                          return await innerDc.endDialog('End Dialog');





                //return await innerDc.beginDialog(SEARCH_REPORT_DIALOG);
                break;
              case 'Weather':

              this.state.cityTemp = ''
              this.state.cityTempHi = ''
              this.state.cityTempLo = ''
              this.state.cityName = ''


                if (!dispatchResults.entities.Cities || dispatchResults.entities.Cities.length === 0 || !dispatchResults.entities.Cities[0]){

                  const cityName = 'Sacramento';

                  var self = this;

                  await axios.get('https://community-open-weather-map.p.rapidapi.com/weather',
                        { params: {
                          'q': String(cityName),
                          'units': 'imperial'
                          },
                        headers: {
                          'X-RapidAPI-Host': process.env.XRapidAPIHost,
                          'X-RapidAPI-Key': process.env.XRapidAPIKey
                    }

                    }).then(response => {

                      if (response){
                        //console.log(response.data)

                        self.state.cityTemp = response.data.main.temp.toFixed(0)
                        self.state.cityTempHi = response.data.main.temp_max.toFixed(0)
                        self.state.cityTempLo = response.data.main.temp_min.toFixed(0)
                        self.state.cityName = response.data.name

                     }

                    }).catch((error)=>{
                           console.log(error);
                    });

                    //Use of Date.now() function
                    var d = Date(Date.now());
                    // Converting the number of millisecond in date string
                    var dateTime = d.toString()

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createWeatherCard(self.state.cityName, dateTime, self.state.cityTemp,self.state.cityTempHi,self.state.cityTempLo)] });

                }else{
                  //console.log(dispatchResults.entities.Cities[0])
                  const cityName = dispatchResults.entities.Cities[0];

                  var self = this;

                  await axios.get('https://community-open-weather-map.p.rapidapi.com/weather',
                        { params: {
                          'q': String(cityName),
                          'units': 'imperial'
                          },
                        headers: {
                          'X-RapidAPI-Host': process.env.XRapidAPIHost,
                          'X-RapidAPI-Key': process.env.XRapidAPIKey
                    }

                    }).then(response => {

                      if (response){
                        //console.log(response.data)

                        self.state.cityTemp = response.data.main.temp.toFixed(0)
                        self.state.cityTempHi = response.data.main.temp_max.toFixed(0)
                        self.state.cityTempLo = response.data.main.temp_min.toFixed(0)
                        self.state.cityName = response.data.name


                     }

                    }).catch((error)=>{
                           console.log(error);
                    });

                    //Use of Date.now() function
                    var d = Date(Date.now());
                    // Converting the number of millisecond in date string
                    var dateTime = d.toString()

                    await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createWeatherCard(self.state.cityName, dateTime, self.state.cityTemp,self.state.cityTempHi,self.state.cityTempLo)] });


                }

                //console.log(dispatchTopIntent)
                break;
              case 'Sports':
                //console.log(dispatchResults.intents)
                //console.log(dispatchResults.entities)
                const teamName = dispatchResults.entities.Sports_Teams[0].toLowerCase();

                var self = this;

                self.state.teamId = ''
                self.state.homeTeam = ''
                self.state.homeTeamBadge = ''
                self.state.homeTeamId = ''
                self.state.homeScore = ''
                self.state.awayTeam = ''
                self.state.awayTeamBadge = ''
                self.state.awayTeamId = ''
                self.state.awayScore = ''
                self.state.dateEvent = ''

                await axios.get('https://www.thesportsdb.com/api/v1/json/1/search_all_teams.php?l=nfl').then(response => {

                    if (response){

                    var itemCount = response.data.teams.length

                    for (var i = 0; i < itemCount; i++)
                    {
                      var teamLowercase = response.data.teams[i].strTeam.toLowerCase()
                      if(teamLowercase.indexOf(teamName) !== -1){
                        //console.log(response.data.teams[i].strTeam.includes(teamName))
                        self.state.teamId = response.data.teams[i].idTeam
                      }

                    }


                   }

                  }).catch((error)=>{
                         console.log(error);
                  });

                  await axios.get('https://www.thesportsdb.com/api/v1/json/1/eventslast.php?id='+self.state.teamId).then(response => {

                      if (response){

                        //console.log(response.data.results[0])

                        self.state.homeTeam = response.data.results[0].strHomeTeam
                        self.state.homeTeamId = response.data.results[0].idHomeTeam
                        self.state.homeScore = response.data.results[0].intHomeScore
                        self.state.awayTeam = response.data.results[0].strAwayTeam
                        self.state.awayTeamId = response.data.results[0].idAwayTeam
                        self.state.awayScore = response.data.results[0].intAwayScore
                        self.state.dateEvent = response.data.results[0].dateEvent


                     }

                    }).catch((error)=>{
                           console.log(error);
                    });

                    await axios.get('https://www.thesportsdb.com/api/v1/json/1/lookupteam.php?id='+self.state.homeTeamId).then(response => {

                        if (response){

                          self.state.homeTeamBadge = response.data.teams[0].strTeamBadge

                       }

                      }).catch((error)=>{
                             console.log(error);
                      });

                      await axios.get('https://www.thesportsdb.com/api/v1/json/1/lookupteam.php?id='+self.state.awayTeamId).then(response => {

                          if (response){

                            self.state.awayTeamBadge = response.data.teams[0].strTeamBadge

                         }

                        }).catch((error)=>{
                               console.log(error);
                        });


                await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createSportCard(self.state.dateEvent, self.state.homeTeam,self.state.homeScore,self.state.homeTeamBadge, self.state.awayTeam,self.state.awayScore, self.state.awayTeamBadge)] });

                break;
              case 'Log_In_As_Guest':
                  return await innerDc.beginDialog(GUEST_LOG_IN_DIALOG);
                  break;
              case 'Log_In':
                // return await innerDc.beginDialog(OAUTH_PROMPT);
                // break;

            }


            switch (dispatchResults.text) {

              case 'ACTO':
                await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
                  //return await innerDc.beginDialog(OAUTH_PROMPT);
                break;
              case 'FINO':
                await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
                // return await step.endDialog();
                break;
              case 'Member':
                await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
                //return await innerDc.endDialog();
                break;
              case 'Employer':
                await innerDc.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('This session is complete, please refresh the page to restart this session','')] });
                //return await innerDc.endDialog();
                break;

            }


        }
    }
}

module.exports.LogoutDialog = LogoutDialog;
