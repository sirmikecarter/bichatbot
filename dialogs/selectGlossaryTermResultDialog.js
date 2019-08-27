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
    async onTurn(turnContext, tokenResponse) {
        // Call QnA Maker and get results.
        //console.log(turnContext.activity.value.report_name_selector_value)

        var glossaryTermQuery = "'" + turnContext.activity.value.glossary_term_selector_value + "'"

        const client = new SimpleGraphClient(tokenResponse.token);
        const me = await client.getMe();

        const definedByToken = me.jobTitle.toLowerCase()

        var self = this;

        await axios.get(process.env.GlossarySearchService +'/indexes/'+ process.env.GlossarySearchServiceIndex + '/docs?',
                { params: {
                  'api-version': '2019-05-06',
                  'search': glossaryTermQuery,
                  '$filter': 'metadata_definedby eq ' + '\'' + definedByToken + '\''
                  },
                headers: {
                  'api-key': process.env.GlossarySearchServiceKey,
                  'ContentType': 'application/json'
          }

        }).then(response => {

          if (response){

          self.state.glossaryTerm= response.data.value[0].questions[0]
          self.state.glossaryDescription = response.data.value[0].answer
          self.state.glossaryDefinedBy = response.data.value[0].metadata_definedby.toUpperCase()
          self.state.glossaryOutput = response.data.value[0].metadata_output.toUpperCase()

         }

        }).catch((error)=>{
               console.log(error);
        });

        await turnContext.sendActivity({ attachments: [this.dialogHelper.createGlossaryCard(me.jobTitle, this.state.glossaryTerm, this.state.glossaryDescription, this.state.glossaryDefinedBy, this.state.glossaryOutput)] });

        await turnContext.sendActivity({ attachments: [this.dialogHelper.createBotCard('...Is there anything else I can help you with?','')] });

        var reply = MessageFactory.suggestedActions(['Main Menu', 'Logout']);
        return await turnContext.sendActivity(reply);

        //return await turnContext.endDialog('End Dialog');

    }
};

module.exports.SelectGlossaryTermResultDialog = SelectGlossaryTermResultDialog;
