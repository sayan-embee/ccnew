// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";

import { Loader, Button, Text, List, Image, Flex } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";

import * as AdaptiveCards from "adaptivecards";

import { getDraftNotification, getConsentSummaries, sendDraftNotification } from '../../apis/messageListApi';

import { getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary, setCardAuthor, setCardBtns, setCardTeamsName } from '../AdaptiveCard/adaptiveCard';
import { getInitAdaptiveCardPDFUpload, setCardTitlePDFUpload, setCardImageLinkPDFUpload, setCardPdfNamePDFUpload, setCardSummaryPDFUpload, setCardAuthorPDFUpload, setCardBtnsPDFUpload, setCardTeamsNamePDFUpload } from '../AdaptiveCard/adaptiveCardPDFUpload';
import { getInitAdaptiveCardQuestionAnswer, setCardTitleQuestionAnswer, setCardAuthorQuestionAnswer, setCardPartQuestionAnswer, setCardBtnNotificationIdQuestionAnswer, setCardBtnQuestionsQuestionAnswer, setCardBtnTitleQuestionAnswer, setCardBtnAuthorQuestionAnswer, setCardTeamsNameQuestionAnswere } from '../AdaptiveCard/adaptiveCardQuestionAnswer';
import { getInitAdaptiveCardEmailTemplate, setCardEmailTemplate } from '../AdaptiveCard/adaptiveCardEmailTemplate';

import { ImageUtil } from '../../utility/imageutility';
import { getBaseUrl } from '../../configVariables';
import { TFunction } from "i18next";

import './sendConfirmationTaskModule.scss';

const pdfImgUrl = getBaseUrl() + "/image/pdfImage.png";

export interface IListItem {
    header: string,
    media: JSX.Element,
}

export interface IMessage {
    id: string;
    title: string;
    acknowledgements?: number;
    reactions?: number;
    responses?: number;
    succeeded?: number;
    failed?: number;
    throttled?: number;
    sentDate?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    buttons: string;
    isImportant?: boolean;
    TemplateType?: any;
    SendTypeId?: any;
    TenantId?: any;
    AdaptiveCardContent?: string;
    emailTitle?: string;
    EmailBody?: any;
    CsvLink?: any;
    AdditionalFileLink?: any,
    tenantName?: string,
    AuthorTeamId?: any,
    authorTeamName?: string,
    AuthorChannelId?: any,
    authorChannelName?: string,
    emailBody?:any;
}


export interface SendConfirmationTaskModuleProps extends RouteComponentProps, WithTranslation {
}

export interface IStatusState {
    message: IMessage;
    loader: boolean;
    teamNames: string[];
    rosterNames: string[];
    groupNames: string[];
    allUsers: boolean;
    messageId?: any;
    templateType?: any,
    questionAnswer: any[],
    sendTypeId: string,
    csvLink?: any

}

class SendConfirmationTaskModule extends React.Component<SendConfirmationTaskModuleProps, IStatusState> {
    readonly localize: TFunction;
    private initMessage = {
        id: "",
        title: "",
        buttons: "[]"
    };

    private card: any;

    constructor(props: SendConfirmationTaskModuleProps) {
        super(props);
        this.localize = this.props.t;

        this.state = {
            message: this.initMessage,
            loader: true,
            teamNames: [],
            rosterNames: [],
            groupNames: [],
            allUsers: false,
            messageId: 0,
            questionAnswer: [],
            sendTypeId: "",
            csvLink: ""
        };
    }

    public componentDidMount() {
        microsoftTeams.initialize();

        let params = this.props.match.params;

        if ('id' in params) {
            let id = params['id'];
            this.getItem(id).then(() => {
                getConsentSummaries(id).then((response) => {
                    this.setState({
                        teamNames: response.data.teamNames.sort(),
                        rosterNames: response.data.rosterNames.sort(),
                        groupNames: response.data.groupNames.sort(),
                        allUsers: response.data.allUsers,
                        messageId: id,
                    }, () => {
                        this.setState({
                            loader: false,
                        }, () => {
                            this.card = (this.state.templateType === this.localize("ImageUpload")) ? getInitAdaptiveCard(this.localize) : (this.state.templateType === this.localize("PDFUpload")) ? getInitAdaptiveCardPDFUpload(this.localize) : (this.state.templateType === this.localize("Q&AUpload")) ? getInitAdaptiveCardQuestionAnswer(this.localize) : getInitAdaptiveCardEmailTemplate(this.localize);

                            // set card properties
                            if (this.state.templateType === this.localize("ImageUpload")) {
                                let teamsName = this.state.message.authorTeamName + " >> " + this.state.message.authorChannelName
                                setCardTeamsName(this.card, teamsName)
                                setCardTitle(this.card, this.state.message.title);
                                setCardImageLink(this.card, this.state.message.imageLink);
                                setCardSummary(this.card, this.state.message.summary);
                                setCardAuthor(this.card, this.state.message.author);
                            }
                            else if (this.state.templateType === this.localize("PDFUpload")) {
                                let teamsName = this.state.message.authorTeamName + " >> " + this.state.message.authorChannelName
                                setCardTeamsNamePDFUpload(this.card, teamsName)
                                setCardTitlePDFUpload(this.card, this.state.message.title);
                                setCardSummaryPDFUpload(this.card, this.state.message.summary);
                                setCardAuthorPDFUpload(this.card, this.state.message.author);
                                // setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                                if (this.state.message.imageLink !== "") {
                                    setCardImageLinkPDFUpload(this.card, pdfImgUrl);

                                    let pdfLink = "[View PDF](" + this.state.message.imageLink + ")"
                                    setCardPdfNamePDFUpload(this.card, pdfLink)

                                }

                            }
                            else if (this.state.templateType === this.localize("Q&AUpload")) {
                                let teamsName = this.state.message.authorTeamName + " >> " + this.state.message.authorChannelName
                                setCardTeamsNameQuestionAnswere(this.card, teamsName)
                                setCardTitleQuestionAnswer(this.card, this.state.message.title);
                                setCardAuthorQuestionAnswer(this.card, this.state.message.author);
                                setCardBtnNotificationIdQuestionAnswer(this.card, this.state.message.id)
                                setCardBtnQuestionsQuestionAnswer(this.card, this.state.questionAnswer)
                                setCardPartQuestionAnswer(this.card, this.state.questionAnswer, this.localize, this.state.message.title, this.state.message.author, teamsName); //update the adaptive cards
                                setCardBtnTitleQuestionAnswer(this.card, this.state.message.title)
                                setCardBtnAuthorQuestionAnswer(this.card, this.state.message.author)

                            }
                            else {
                                
                                if (this.state.message.imageLink !== "") {
                                    let teamsName = this.state.message.authorTeamName + " >> " + this.state.message.authorChannelName
                                    let emailLink = "[" + this.state.message.emailTitle + "](" + this.state.message.imageLink + ")"
                                    setCardEmailTemplate(this.card, teamsName, this.state.message.emailBody, this.state.message.title, this.state.message.author, emailLink, this.state.message.summary)  
                                }
                            }


                            if (this.state.message.buttonTitle && this.state.message.buttonLink && !this.state.message.buttons) {
                                setCardBtns(this.card, [{
                                    "type": "Action.OpenUrl",
                                    "title": this.state.message.buttonTitle,
                                    "url": this.state.message.buttonLink
                                }]);


                            }
                            else {
                                if (this.state.templateType === this.localize("ImageUpload")) {
                                    setCardBtns(this.card, JSON.parse(this.state.message.buttons));
                                }
                                else if (this.state.templateType === this.localize("PDFUpload")) {
                                    setCardBtnsPDFUpload(this.card, JSON.parse(this.state.message.buttons));
                                }
                            }

                            let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                            adaptiveCard.parse(this.card);
                            let renderedCard = adaptiveCard.render();
                            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                            if (this.state.message.buttonLink) {
                                let link = this.state.message.buttonLink;
                                adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                            }
                        });
                    });
                });
            });
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getDraftNotification(id);
            console.log("send get item", response.data)
            if (response.data.templateType === this.localize("Q&AUpload")) {
                if (response.data.summary !== "") {
                    this.setState({
                        questionAnswer: JSON.parse(response.data.summary)
                    });
                }
            }

            this.setState({
                message: response.data,
                templateType: response.data.templateType,
                sendTypeId: response.data.sendTypeId,
                csvLink: response.data.csvLink
            });
        } catch (error) {
            return error;
        }
    }

    public render(): JSX.Element {
        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            return (
                <div className="taskModule">
                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                        <Flex className="scrollableContent" gap="gap.small">
                            <Flex.Item size="size.half">
                                <Flex column className="formContentContainer">
                                    <h3>{this.localize("ConfirmToSend")}</h3>
                                    <span>{this.localize("SendToRecipientsLabel")}</span>

                                    <div className="results">
                                        {this.renderAudienceSelection()}
                                    </div>

                                    <h3>{this.localize("Tenant")}</h3>
                                    <label>{this.state.message.tenantName}</label>

                                    <h3>{this.localize("Important")}</h3>
                                    <label>{this.renderImportant()}</label>
                                </Flex>
                            </Flex.Item>
                            <Flex.Item size="size.half">
                                <div className="adaptiveCardContainer">
                                </div>
                            </Flex.Item>
                        </Flex>
                        <Flex className="footerContainer" vAlign="end" hAlign="end">
                            <Flex className="buttonContainer" gap="gap.small">
                                <Flex.Item push>
                                    <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={this.localize("PreparingMessageLabel")} labelPosition="end" />
                                </Flex.Item>
                                <Button content={this.localize("Send")} id="sendBtn" onClick={this.onSendMessage} primary />
                            </Flex>
                        </Flex>
                    </Flex>
                </div>
            );
        }
    }

    private onSendMessage = () => {
        let spanner = document.getElementsByClassName("sendingLoader");
        spanner[0].classList.remove("hiddenLoader");

        this.setState({
            message: { ...this.state.message, AdaptiveCardContent: JSON.stringify(this.card) }
        }, () => {
            console.log("adaptive card", this.state.message)
            sendDraftNotification(this.state.message).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        })


    }

    private getItemList = (items: string[]) => {
        let resultedTeams: IListItem[] = [];
        if (items) {
            resultedTeams = items.map((element) => {
                const resultedTeam: IListItem = {
                    header: element,
                    media: <Image src={ImageUtil.makeInitialImage(element)} avatar />
                }
                return resultedTeam;
            });
        }
        return resultedTeams;
    }

    private renderImportant = () => {
        if (this.state.message.isImportant) {
            return (
                <label>Yes</label>
            )
        } else {
            return (
                <label>No</label>
            )
        }
    }

    private renderAudienceSelection = () => {
        if (this.state.teamNames && this.state.teamNames.length > 0) {
            return (
                <div key="teamNames"> <span className="label">{this.localize("TeamsLabel")}</span>
                    <List items={this.getItemList(this.state.teamNames)} />
                </div>
            );
        } else if (this.state.rosterNames && this.state.rosterNames.length > 0) {
            return (
                <div key="rosterNames"> <span className="label">{this.localize("TeamsMembersLabel")}</span>
                    <List items={this.getItemList(this.state.rosterNames)} />
                </div>);
        } else if (this.state.groupNames && this.state.groupNames.length > 0) {
            return (
                <div key="groupNames" > <span className="label">{this.localize("GroupsMembersLabel")}</span>
                    <List items={this.getItemList(this.state.groupNames)} />
                </div>);
        } else if (this.state.allUsers) {
            return (
                <div key="allUsers">
                    <span className="label">{(this.state.sendTypeId === "5") ? this.localize("AllUsersSisterTenantLabel") : this.localize("AllUsersLabel")}</span>
                    <div className="noteText">
                        <Text error content={(this.state.sendTypeId === "5") ? this.localize("SendToAllUsersSisterTenantNote") : this.localize("SendToAllUsersNote")} />
                    </div>
                </div>);
        } else if (this.state.csvLink && this.state.sendTypeId === "6") {
            return (<div className="flexColumn">
                <Text error content={this.localize("SendToCSVFileNote")} />
                <a href={this.state.csvLink}>{this.localize("ViewFile")}</a>
            </div>);
        }
        else {
            return (<div></div>);
        }
    }
}

const sendConfirmationTaskModuleWithTranslation = withTranslation()(SendConfirmationTaskModule);
export default sendConfirmationTaskModuleWithTranslation;
