const { ComponentDialog, ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const { DialogHelper } = require('./dialogHelper');
const axios = require('axios');
var arraySort = require('array-sort');

const CONFIRM_PROMPT = 'confirmPrompt';
const CHOICE_PROMPT = 'CHOICE_PROMPT';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class SearchWeatherDialog extends ComponentDialog {
    constructor(id) {
        super(id || 'searchWeatherDialog');

        this.dialogHelper = new DialogHelper();

        this.state = {
          cityName: '',
          cityTemp: '',
          cityTempHi: '',
          cityTempLo: ''
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

      this.state.cityTemp = ''
      this.state.cityTempHi = ''
      this.state.cityTempLo = ''
      this.state.cityName = ''

      const cityName = stepContext._info.options.city_name;

      console.log(cityName)

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
                   //console.log(error);
            });

            //Use of Date.now() function
            var d = Date(Date.now());
            // Converting the number of millisecond in date string
            var dateTime = d.toString()

            if(self.state.cityName){
              await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createWeatherCard(self.state.cityName, dateTime, self.state.cityTemp,self.state.cityTempHi,self.state.cityTempLo)] });
            }else{
              await stepContext.context.sendActivity({ attachments: [this.dialogHelper.createBotCard('...City Name Not Found','')] });
            }




      return await stepContext.endDialog('End Dialog');
    }

    async resultStep(stepContext) {

      return await stepContext.endDialog('End Dialog');

    }

}

module.exports.SearchWeatherDialog = SearchWeatherDialog;
