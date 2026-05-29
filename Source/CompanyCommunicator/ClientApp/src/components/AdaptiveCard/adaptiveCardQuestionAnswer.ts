// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TFunction } from "i18next";
import * as AdaptiveCards from "adaptivecards";
import MarkdownIt from "markdown-it";
import { getBaseUrl } from './../../configVariables';

// Static method to render markdown on the adaptive card
AdaptiveCards.AdaptiveCard.onProcessMarkdown = function (text, result) {
    var md = new MarkdownIt();
    // Teams only supports a subset of markdown as per https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-format?tabs=adaptive-md%2Cconnector-html#formatting-cards-with-markdown
    md.disable(['image', 'table', 'heading',
        'hr', 'code', 'reference',
        'lheading', 'html_block', 'fence',
        'blockquote', 'strikethrough']);
    // renders the text
    result.outputHtml = md.render(text);
    result.didProcess = true;
}

const  baseUrl = getBaseUrl() ;

export const getInitAdaptiveCardQuestionAnswer = (t: TFunction) => {
    const titleTextAsString = t("TitleText");
    return (
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "size": "Large",
                    "weight": "Bolder",
                    "text": "",
                    "color": "Attention"
                },
                {
                    "type": "TextBlock",
                    "weight": "Bolder",
                    "text": titleTextAsString,
                    "size": "ExtraLarge",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": ""
                },
                {
                    "type": "Image",
                    "url": baseUrl + "/api/ReadReceiptNotification/view",
                    "size": "Small",
                    "width": "1px",
                    "height": "1px"
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit",
                    "data":{
                        "NotificationId":"",
                        "Questions":"",
                        "Title":"",
                        "Author":""
                    }  
                }
            ],
            "msteams": {
                "width": "Full"
            },
            "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.4"
        }
    );

}

export const getCardTeamsNameQuestionAnswer = (card: any) => {
    return card.body[0].text;
}

export const setCardTeamsNameQuestionAnswere = (card: any, teamsName: string) => {
    card.body[0].text = teamsName;
}

export const getCardTitleQuestionAnswer = (card: any) => {
    return card.body[1].text;
}

export const setCardTitleQuestionAnswer = (card: any, title: string) => {
    card.body[1].text = title;
}

export const getCardAuthorQuestionAnswer = (card: any) => {
    return card.body[2].text;
}

export const setCardAuthorQuestionAnswer = (card: any, author?: string) => {
    card.body[2].text = author;
}

export const getCardBtnNotificationIdQuestionAnswer = (card: any) => {
    return card.actions[0].data.NotificationId;
}

export const setCardBtnNotificationIdQuestionAnswer = (card: any, id?:any) => {
    card.actions[0].data.NotificationId = id;
}

export const setCardBtnQuestionsQuestionAnswer = (card: any, questionSet: any[]) => {
   const question= questionSet.map((a:any) => `${a.question}`).join('||');
   card.actions[0].data.Questions=question
}

export const getCardBtnQuestionsQuestionAnswer = (card: any) => {
     return card.actions[0].data.Questions;
 }

 export const getCardBtnTitleQuestionAnswer = (card: any) => {
    return card.actions[0].data.Title;
}

export const setCardBtnTitleQuestionAnswer = (card: any, title?:string) => {
    card.actions[0].data.Title = title;
}

export const getCardBtnAuthorQuestionAnswer = (card: any) => {
    return card.actions[0].data.Author;
}

export const setCardBtnAuthorQuestionAnswer = (card: any, author?:any) => {
    card.actions[0].data.Author = author;
}



export const setCardPartQuestionAnswer = (card: any, questionSet: any[], localize: any, title: any, author: any, teamsName: any) => {
    // console.log("Question answer", questionSet);
    var jsonArrayQuestions = new Array();
    jsonArrayQuestions.push(`{
        "type": "TextBlock",
        "size": "Large",
        "weight": "Bolder",
        "text": "${teamsName}",
        "color": "Attention"
    }`)
    jsonArrayQuestions.push(`{
            "type": "TextBlock",
            "weight": "Bolder",
            "text": "${title}",
            "size": "ExtraLarge",
            "wrap": true
        }`)
    jsonArrayQuestions.push(`{
            "type": "TextBlock",
            "wrap": true,
            "size": "Small",
            "weight": "Lighter",
            "text": "${author}"
        }`)
        jsonArrayQuestions.push(`{
            "type": "Image",
            "url": "${baseUrl}/api/ReadReceiptNotification/view",
            "size": "Small",
            "width": "1px",
            "height": "1px"
        }`)
    if (questionSet !== null) {
        questionSet.forEach((item,index) => {
            if (item.questionType === localize("DescriptiveQuestion")) {
                jsonArrayQuestions.push(`{
                            "type": "TextBlock",
                            "text": "${item.question}",
                            "wrap": true,
                            "size": "Medium",
                            "id":  "question${index}"
                        }`);
                jsonArrayQuestions.push(`{
                            "type": "Input.Text",
                            "placeholder": "${localize("Answer")}",
                            "wrap": true,
                            "id":  "answer${index}",
                            "isRequired": true,
                            "errorMessage": "Please enter your input."
                        }`);
            }
            else {
                jsonArrayQuestions.push(`{
                            "type": "TextBlock",
                            "text": "${item.question}",
                            "wrap": true,
                            "id":  "question${index}"
                        }`);
                 
                    jsonArrayQuestions.push(`{
                            "type": "Input.ChoiceSet",
                            "style": "expanded",
                            "id":  "answer${index}",
                            "isRequired": true,
                            "errorMessage": "Please enter your input.",
                            "choices": [
                                {  
                                    "title": "${item.answer[0] ? item.answer[0]:""}",
                                    "value": "${item.answer[0] ? item.answer[0]:""}"
                                },
                                {
                                    "title": "${item.answer[1] ? item.answer[1]:""}",
                                    "value": "${item.answer[1] ? item.answer[1]:""}"
                                },
                                {
                                    "title": "${item.answer[2] ? item.answer[2]:""}",
                                    "value": "${item.answer[2] ? item.answer[2]:""}"
                                },
                                {
                                    "title": "${item.answer[3] ? item.answer[3]:""}",
                                    "value": "${item.answer[3] ? item.answer[3]:""}"
                                }
                            ]
                        }`);
            }

        });




        var jsonQuestionCard = "[" + jsonArrayQuestions.join() + "]";
        card.body = JSON.parse(jsonQuestionCard);
        return card.body[0];

    } else {
    }
}


