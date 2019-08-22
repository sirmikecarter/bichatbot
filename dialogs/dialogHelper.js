// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { AttachmentLayoutTypes, CardFactory, MessageFactory } = require('botbuilder-core');

class DialogHelper {

     createMenu(title,actionTitle) {
       return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "TextBlock",
             "text": title,
             "weight": "bolder",
             "size": "medium"
           }],
         "actions": [
           {
             "type": "Action.Submit",
             "title": actionTitle,
             "data": 'luis: '+ title + ' ' + actionTitle
           }
         ]
       });
     }

     createGifCard() {

       return CardFactory.animationCard(
           '2%',
           [
               { url: 'http://i.imgur.com/ptJ6Ph6.gif' }
           ],
           [],
           {
               subtitle: 'Retirement Formula'
           }
       );
     }

     createDocumentCard(title, language, keyPhrases, organizations, persons, locations, glossary1, glossary2) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "TextBlock",
             "text": title,
             "weight": "bolder",
             "size": "medium"
           }
         ],
         "actions": [
           {
             "type": "Action.ShowCard",
             "title": "Language",
             "card": {
               "type": "AdaptiveCard",
               "body": [
                 {
                   "type": "TextBlock",
                   "text": "Document Language:",
                   "weight": "bolder",
                   "size": "small",
                   "separator": true
                 },
                 {
                   "type": "TextBlock",
                   "text": language,
                   "size": "small",
                   "wrap": true
                 },
               ]
             }
           },
             {
               "type": "Action.ShowCard",
               "title": "Key Phrases",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Key Phrases:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": keyPhrases + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Organizations",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Organizations",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": organizations + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Persons",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Persons",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": persons + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Locations",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Locations",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": locations + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "CalPERS Glossary",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "CalPERS Glossary",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": glossary1 + "\r",
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": glossary2 + "\r",
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             }
           ]
       });
     }

     createReportCard(title, description, owner, designee, approver, division, classification, language, entities, keyPhrases, sentiment) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "TextBlock",
             "text": title,
             "weight": "bolder",
             "size": "medium"
           },
           {
             "type": "TextBlock",
             "text": description,
             "wrap": true
           }
         ],
         "actions": [
             {
               "type": "Action.ShowCard",
               "title": "Additional Information",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Additional Information",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "FactSet",
                     "facts": [
                       {
                         "title": "Owner:",
                         "value": owner,
                         "wrap": true
                       },
                       {
                         "title": "Designee:",
                         "value": designee,
                         "wrap": true
                       },
                       {
                         "title": "Approver:",
                         "value": approver,
                         "wrap": true
                       },
                       {
                         "title": "Division:",
                         "value": division,
                         "wrap": true
                       },
                       {
                         "title": "Classification:",
                         "value": classification,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "Text Analytics",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                   {
                     "type": "TextBlock",
                     "text": "Text Analytics",
                     "weight": "bolder",
                     "size": "medium",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Report Language:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": language,
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Entities:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": entities + "\r",
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Key Phrases:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": keyPhrases + "\r",
                     "size": "small",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "text": "Sentiment Score:",
                     "weight": "bolder",
                     "size": "small",
                     "separator": true
                   },
                   {
                     "type": "TextBlock",
                     "text": sentiment,
                     "size": "small",
                     "wrap": true
                   }
                 ]
               }
             },
             {
               "type": "Action.OpenUrl",
               "title": "View Report",
               "url": "http://adaptivecards.io"
             }
           ]
       });
     }

     createBotCard(text1, text2) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "ColumnSet",
             "columns": [
               {
                 "type": "Column",
                 "width": "auto",
                 "items": [
                   {
                     "type": "Image",
                     "url": "https://gateway.ipfs.io/ipfs/QmXKfQgKVckfbGSMmzHAGAZ3zr1h8yJNrmEuBaJdNsGECs",
                     "size": "small",
                     "style": "person"
                   }
                 ]
               },
               {
                 "type": "Column",
                 "width": "stretch",
                 "items": [
                   {
                     "type": "TextBlock",
                     "text": text1,
                     "weight": "bolder",
                     "wrap": true
                   },
                   {
                     "type": "TextBlock",
                     "spacing": "none",
                     "text": text2,
                     "isSubtle": true,
                     "wrap": true
                   }
                 ]
               }
             ]
           }
         ]
       });
     }

     createComboListCard(choiceList, selectorValue) {

     return CardFactory.adaptiveCard({
       "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
       "type": "AdaptiveCard",
       "version": "1.0",
       "body": [
         {
           "type": "Input.ChoiceSet",
           "id": selectorValue,
           "style": "compact",
           "value": "0",
           "choices": choiceList
         }
       ],
       "actions": [
         {
           "type": "Action.Submit",
           "id": "submit",
           "title": "Submit",
           "data":{
                 "action": selectorValue
           }
         }
       ]
     });
     }
}

module.exports.DialogHelper = DialogHelper;
