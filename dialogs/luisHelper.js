// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { LuisRecognizer } = require('botbuilder-ai');

class LuisHelper {
    /**
     * Returns an object with preformatted LUIS results for the bot's dialogs to consume.
     * @param {*} logger
     * @param {TurnContext} context
     */
    static async executeLuisQuery(logger, context) {
        const bookingDetails = {};

        try {
            const recognizer = new LuisRecognizer({
                applicationId: process.env.LuisAppId,
                endpointKey: process.env.LuisAPIKey,
                endpoint: `https://${ process.env.LuisAPIHostName }`
            }, {}, true);

            const recognizerResult = await recognizer.recognize(context);


            //console.log(recognizerResult.entities['Report_Classification'])

            const intent = LuisRecognizer.topIntent(recognizerResult);

            bookingDetails.intent = intent;

            if (intent === 'Search') {
                // We need to get the result from the LUIS JSON which at every level returns an array

                bookingDetails.reportname = recognizerResult.entities['Report_Name']
                bookingDetails.description = recognizerResult.entities['Report_Description']
                bookingDetails.owner = recognizerResult.entities['Report_Owner']
                bookingDetails.designee = recognizerResult.entities['Report_Designee']
                bookingDetails.approver = recognizerResult.entities['Report_Approver']
                bookingDetails.division= recognizerResult.entities['Report_Division']
                bookingDetails.classification = recognizerResult.entities['Report_Classification']

                //bookingDetails.destination = LuisHelper.parseCompositeEntity(recognizerResult, 'To', 'Airport');
                //bookingDetails.origin = LuisHelper.parseCompositeEntity(recognizerResult, 'From', 'Airport');

                // This value will be a TIMEX. And we are only interested in a Date so grab the first result and drop the Time part.
                // TIMEX is a format that represents DateTime expressions that include some ambiguity. e.g. missing a Year.
                //bookingDetails.travelDate = LuisHelper.parseDatetimeEntity(recognizerResult);
            }
        } catch (err) {
            logger.warn(`LUIS Exception: ${ err } Check your LUIS configuration`);
        }
        return bookingDetails;
    }

    static parseCompositeEntity(result, compositeName, entityName) {
        const compositeEntity = result.entities[compositeName];
        if (!compositeEntity || !compositeEntity[0]) return undefined;

        const entity = compositeEntity[0][entityName];
        if (!entity || !entity[0]) return undefined;

        const entityValue = entity[0][0];
        return entityValue;
    }

    static parseDatetimeEntity(result) {
        const datetimeEntity = result.entities['datetime'];
        if (!datetimeEntity || !datetimeEntity[0]) return undefined;

        const timex = datetimeEntity[0]['timex'];
        if (!timex || !timex[0]) return undefined;

        const datetime = timex[0].split('T')[0];
        return datetime;
    }
}

module.exports.LuisHelper = LuisHelper;
