const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./helpers/dialogHelper');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

const axios = require('axios');

class SelectSportsTeamDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'selectSportsTeamDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          teamName1: '',
          teamName2: '',
          teamId: '',
          teamIdMLB: '',
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
      self.state.teamName1 = ''
      self.state.teamName2 = ''

      console.log('TEAM1: ' + stepContext._info.options.team1)
      console.log('TEAM2: ' + stepContext._info.options.team2)

      await axios.get('https://www.thesportsdb.com/api/v1/json/1/lookupteam.php?id='+stepContext._info.options.team1).then(response => {

          if (response){

          self.state.teamName1 = response.data.teams[0].strTeam

         }

        }).catch((error)=>{
               console.log(error);
        });

        await axios.get('https://www.thesportsdb.com/api/v1/json/1/lookupteam.php?id='+stepContext._info.options.team2).then(response => {

            if (response){

            self.state.teamName2 = response.data.teams[0].strTeam

           }

          }).catch((error)=>{
                 console.log(error);
          });

          await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...There are two teams with the same Name, please select the correct one','')] });

          return await stepContext.prompt(CHOICE_PROMPT, {
              prompt: '',
              choices: ChoiceFactory.toChoices([self.state.teamName1, self.state.teamName2])
          });


      // await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });
      //
      // var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
      // await stepContext.context.sendActivity(reply);

      //return await stepContext.endDialog('End Dialog');
    }

    async resultStep(stepContext) {

      //console.log(stepContext.result.value)

      const teamName = stepContext.result.value.toLowerCase();

      var self = this;

      self.state.teamId = ''
      self.state.teamIdMLB = ''
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

        await axios.get('https://www.thesportsdb.com/api/v1/json/1/search_all_teams.php?l=mlb').then(response => {

            if (response){

            var itemCountMLB = response.data.teams.length

            for (var i = 0; i < itemCountMLB; i++)
            {
              var teamLowercaseMLB = response.data.teams[i].strTeam.toLowerCase()
              if(teamLowercaseMLB.indexOf(teamName) !== -1){
                //console.log(response.data.teams[i].strTeam.includes(teamName))
                self.state.teamIdMLB = response.data.teams[i].idTeam
              }

            }


           }

          }).catch((error)=>{
                 console.log(error);
          });

          if(self.state.teamId === ''){
            self.state.teamId = self.state.teamIdMLB
          }

          // if(self.state.teamId !== '' && self.state.teamIdMLB !== '' ){
          //   console.log(self.state.teamId)
          //   console.log(self.state.teamIdMLB)
          //   return await innerDc.beginDialog(SELECT_SPORTS_TEAM_DIALOG, { team1: self.state.teamId, team2: self.state.teamIdMLB});
          //   //return await this.selectSportsTeamDialog.destinationStep(innerDc, self.state.teamId, self.state.teamIdMLB);
          // }

          // console.log(self.state.teamId)
          // console.log(self.state.teamIdMLB)

          if(self.state.teamId){

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


      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createSportCard(self.state.dateEvent, self.state.homeTeam,self.state.homeScore,self.state.homeTeamBadge, self.state.awayTeam,self.state.awayScore, self.state.awayTeamBadge)] });

    }else{

      await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...No Results Found','')] });

    }





      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SelectSportsTeamDialog = SelectSportsTeamDialog;
