// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { ComponentDialog, ChoicePrompt, DialogSet, DialogTurnStatus, OAuthPrompt, TextPrompt, WaterfallDialog, ChoiceFactory } = require('botbuilder-dialogs');
const { LuisApplication, LuisPredictionOptions, LuisRecognizer, QnAMaker } = require('botbuilder-ai');

const { ActivityTypes } = require('botbuilder');
const { DialogHelper } = require('./dialogHelper');
const { GuestLogInDialog } = require('./guestLogInDialog');
const { SearchGlossaryTermDialog } = require('./searchGlossaryTermDialog');
const { SearchReportDialog } = require('./searchReportDialog');
const axios = require('axios');


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
        reportsFilterString: ''
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
                //console.log(dispatchResults)
                //console.log(dispatchResults.entities.Term[0])
                return await innerDc.beginDialog(SEARCH_GLOSSARY_TERM_DIALOG);
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

                      console.log(itemCount)
                      //
                      // var itemArray = self.state.reportArray.slice();
                      //
                      // for (var i = 0; i < itemCount; i++)
                      // {
                      //       const reportname = response.data.value[i].metadata_reportname
                      //       const description = response.data.value[i].answer
                      //       const owner = response.data.value[i].metadata_owner
                      //       const designee = response.data.value[i].metadata_designee
                      //       const approver = response.data.value[i].metadata_approver
                      //       const division = response.data.value[i].metadata_division
                      //       const classification = response.data.value[i].metadata_classification
                      //
                      //       if (itemArray.indexOf(reportname) === -1)
                      //       {
                      //         itemArray.push({'reportname': reportname, 'description': description, 'owner': owner, 'designee': designee, 'approver': approver, 'division': division, 'classification': classification})
                      //       }
                      // }
                      //
                      // self.state.reportArray = itemArray

                   }

                  }).catch((error)=>{
                         console.log(error);
                  });





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
