// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
const { QnAMaker } = require('botbuilder-ai');
const { LuisHelper } = require('./luisHelper');
const { LuisRecognizer } = require('botbuilder-ai');
const { DialogHelper } = require('./dialogHelper');
const { SimpleGraphClient } = require('../simple-graph-client');
const { ConfirmPrompt, TextPrompt, WaterfallDialog, ChoiceFactory, ChoicePrompt, DialogSet } = require('botbuilder-dialogs');
const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');
const axios = require('axios');
var arraySort = require('array-sort');

// Name of the QnA Maker service in the .bot file.
const QNA_CONFIGURATION = 'q_sample-qna';
// CONSTS used in QnA Maker query. See [here](https://docs.microsoft.com/en-us/azure/bot-service/bot-builder-howto-qna?view=azure-bot-service-4.0&tabs=cs) for additional info
const QNA_TOP_N = 1;
const QNA_CONFIDENCE_THRESHOLD = 0.5;


class SelectGlossaryTermResultDialog {
    /**
     *

     */
    constructor() {

        this.dialogHelper = new DialogHelper();


        this.state = {
          termArray: [],
          glossaryTerm: '',
          glossaryDescription: '',
          glossaryDefinedBy: '',
          glossaryOutput: ''
        };
    }

    /**
     *
     * @param {TurnContext} turn context object
     */
    async onTurn(stepContext, turnContext, tokenResponse) {
        // Call QnA Maker and get results.
        //console.log(turnContext.activity.value.report_name_selector_value)



        var glossaryTermQuery = "'" + turnContext.activity.value.glossary_term_selector_value + "'"

        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();

        const definedByToken = me.jobTitle.toLowerCase()

        var self = this;
        self.state.termArray = []

        await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
                { params: {
                  'api-version': '2019-05-06',
                  '$filter': 'search.ismatch(' + glossaryTermQuery + ',' + '\'questions\'' + ')'
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

                    if (turnContext.activity.value.glossary_term_selector_value === glossaryTerm ){
                      itemArray.push({'first': 'a', 'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput, 'related': glossaryRelated})
                    }else{
                      itemArray.push({'first': 'z', 'glossaryterm': glossaryTerm, 'description': glossaryDescription, 'definedby': glossaryDefinedBy, 'output': glossaryOutput, 'related': glossaryRelated})
                    }
                  }

            }

            self.state.termArray = arraySort(itemArray, 'first')


         }

        }).catch((error)=>{
               console.log(error);
        });

        if(definedByToken.toUpperCase() === self.state.termArray[0].definedby){
          await turnContext.sendActivity({ attachments: [this.dialogHelper.createBotCard('...This Term Is Defined By Your Area:','')] });
        }else{
          await turnContext.sendActivity({ attachments: [this.dialogHelper.createBotCard('...This Term Is NOT Defined By Your Area:','')] });
        }

        await turnContext.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(self.state.termArray[0].definedby, self.state.termArray[0].glossaryterm, self.state.termArray[0].description, self.state.termArray[0].definedby, self.state.termArray[0].output, self.state.termArray[0].related)] });



        // console.log(self.state.termArray.length)
        // console.log(self.state.termArray)

        if (self.state.termArray.length > 1){
          await turnContext.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Also Found a Similar Term:','')] });
          await turnContext.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(self.state.termArray[1].definedby, self.state.termArray[1].glossaryterm, self.state.termArray[1].description, self.state.termArray[1].definedby, self.state.termArray[1].output, self.state.termArray[1].related)] });

        }

        await turnContext.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

      return await stepContext.prompt(CHOICE_PROMPT, {
            prompt: '',
            choices: ChoiceFactory.toChoices(['Main Menu', 'Logout'])
        });


        // var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
        // return await turnContext.sendActivity(reply);

        //return await turnContext.endDialog('End Dialog');

    }
};

module.exports.SelectGlossaryTermResultDialog = SelectGlossaryTermResultDialog;
