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

     createLink(text, link) {
       return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
             "type": "TextBlock",
             "text": text
           }
         ],
         "actions": [
           {
             "type": "Action.OpenUrl",
             "title": "Click Me",
             "url": link
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

     createImageCard() {

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
                         "items": [
                             {
                                 "type": "Image",
                                 "url": "https://cdn.dribbble.com/users/334335/screenshots/3026014/halloween-eblast-header-3.png"
                             }
                         ]
                     }
                 ]
             }
         ]
       });
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

     createGlossaryCard(division, term, description, definedBy, output, related) {

     return CardFactory.adaptiveCard({
         "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
         "type": "AdaptiveCard",
         "version": "1.0",
         "body": [
           {
            "type": "TextBlock",
            "text": division + " Business Glossary",
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": term,
             "weight": "bolder",
             "size": "medium",
             "separator": true
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
                         "title": "Defined By:",
                         "value": definedBy,
                         "wrap": true
                       },
                       {
                         "title": "Output:",
                         "value": output,
                         "wrap": true
                       },
                       {
                         "title": "Related Terms:",
                         "value": related,
                         "wrap": true
                       }
                     ]
                   },
                 ]
               }
             },
             {
               "type": "Action.ShowCard",
               "title": "View Mind Map",
               "card": {
                 "type": "AdaptiveCard",
                 "body": [
                          {
                              "type": "TextBlock",
                              "text": "Mind Map",
                              "weight": "Bolder",
                              "horizontalAlignment": "Center",
                              "size": "Large"
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Related",
                                              "wrap": true,
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder"
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Term",
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder"
                                          }
                                      ],
                                      "verticalContentAlignment": "Center",
                                      "horizontalAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Output",
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder"
                                          }
                                      ],
                                      "horizontalAlignment": "Center",
                                      "verticalContentAlignment": "Center"
                                  }
                              ]
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": related,
                                              "horizontalAlignment": "Center",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "Image",
                                              "url": "http://messagecardplayground.azurewebsites.net/assets/graydot2x2.png",
                                              "width": "1000px",
                                              "height": "2px"
                                          }
                                      ],
                                      "horizontalAlignment": "Center",
                                      "spacing": "None",
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": term,
                                              "horizontalAlignment": "Center",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "Image",
                                              "url": "http://messagecardplayground.azurewebsites.net/assets/graydot2x2.png",
                                              "width": "1000px",
                                              "height": "2px"
                                          }
                                      ],
                                      "spacing": "None",
                                      "horizontalAlignment": "Center",
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                          "type": "TextBlock",
                                          "text": output,
                                          "horizontalAlignment": "Center",
                                          "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  }
                              ],
                              "separator": true
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "Image",
                                              "altText": "",
                                              "url": "http://messagecardplayground.azurewebsites.net/assets/graydot2x2.png",
                                              "width": "2px",
                                              "horizontalAlignment": "Center",
                                              "height": "25px"
                                          }
                                      ]
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  }
                              ],
                              "separator": true
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": "Defined By",
                                              "horizontalAlignment": "Center",
                                              "weight": "Bolder",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  }
                              ]
                          },
                          {
                              "type": "ColumnSet",
                              "columns": [
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch",
                                      "items": [
                                          {
                                              "type": "TextBlock",
                                              "text": definedBy,
                                              "horizontalAlignment": "Center",
                                              "wrap": true
                                          }
                                      ],
                                      "verticalContentAlignment": "Center",
                                      "horizontalAlignment": "Center"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  },
                                  {
                                      "type": "Column",
                                      "width": "stretch"
                                  }
                              ]
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
            "text": "Archer Reports",
            "weight": "bolder",
            "isSubtle": false
           },
           {
             "type": "TextBlock",
             "text": title,
             "weight": "bolder",
             "size": "medium",
             "separator": true
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

     createUserCard(picture, name, division) {

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
                     "url": picture,
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
                     "text": name,
                     "weight": "bolder",
                     "wrap": true
                   },
                    {
                      "type": "TextBlock",
                      "spacing": "none",
                      "text": division,
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

     createComboListCard(helperText, choiceList, selectorValue) {

     return CardFactory.adaptiveCard({
       "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
       "type": "AdaptiveCard",
       "version": "1.0",
       "body": [
         {
          "type": "TextBlock",
          "text": helperText,
          "weight": "bolder",
          "isSubtle": false
         },
         {
           "type": "Input.ChoiceSet",
           "id": selectorValue,
           "style": "compact",
           "value": "0",
           "separator": true,
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
