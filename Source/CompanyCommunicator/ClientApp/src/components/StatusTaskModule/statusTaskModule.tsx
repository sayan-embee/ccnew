// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import './statusTaskModule.scss';
import { getSentNotification, exportNotification, surveyexport, reactionexport, getRecallNotification, recallDataExport } from '../../apis/messageListApi';
import { RouteComponentProps } from 'react-router-dom';
import * as AdaptiveCards from "adaptivecards";
import { TooltipHost } from 'office-ui-fabric-react';
import { Loader, List, Image, Button, DownloadIcon, AcceptIcon, Flex, Dialog } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";
import { CSVLink } from "react-csv";

import { getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary, setCardAuthor, setCardBtns, setCardTeamsName } from '../AdaptiveCard/adaptiveCard';
import { getInitAdaptiveCardPDFUpload, setCardTitlePDFUpload, setCardImageLinkPDFUpload, setCardPdfNamePDFUpload, setCardSummaryPDFUpload, setCardAuthorPDFUpload, setCardBtnsPDFUpload, setCardTeamsNamePDFUpload } from '../AdaptiveCard/adaptiveCardPDFUpload';
import { getInitAdaptiveCardQuestionAnswer, setCardTitleQuestionAnswer, setCardAuthorQuestionAnswer, setCardPartQuestionAnswer, setCardTeamsNameQuestionAnswere } from '../AdaptiveCard/adaptiveCardQuestionAnswer';
import { getInitAdaptiveCardEmailTemplate, setCardEmailTemplate } from '../AdaptiveCard/adaptiveCardEmailTemplate';

import { ImageUtil } from '../../utility/imageutility';
import { formatDate, formatDuration, formatNumber } from '../../i18n';
import { TFunction } from "i18next";
import { getBaseUrl } from '../../configVariables';

const pdfImgUrl = getBaseUrl() + "/image/pdfImage.png";


export interface IListItem {
    header: string,
    media: JSX.Element,
}

export interface IMessage {
    id: string;
    title: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
    succeeded?: string;
    failed?: string;
    unknown?: string;
    sentDate?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    teamNames?: string[];
    rosterNames?: string[];
    groupNames?: string[];
    allUsers?: boolean;
    sendingStartedDate?: string;
    sendingDuration?: string;
    errorMessage?: string;
    warningMessage?: string;
    canDownload?: boolean;
    sendingCompleted?: boolean;
    buttons: string;
    isImportant?: boolean;
    emailTitle?: string;
    emailBody?: any;
    CsvLink?: any;
    AdditionalFileLink?: any,
    tenantName?: string,
    AuthorTeamId?: any,
    authorTeamName?: string,
    AuthorChannelId?: any,
    authorChannelName?: string,
    recalledDate?: any
}

export interface IStatusState {
    message: IMessage;
    loader: boolean;
    page: string;
    teamId?: string;
    templateType: string;
    questionAnswer: any[];
    id?: any;
    downloadData?: any;
    adaptiveCardContent?: any;
    totalReadReceipt?: any;
    surveyDownloadData?: any;
    notificationId?: any;
    openFrom?: any;
}

interface StatusTaskModuleProps extends RouteComponentProps, WithTranslation { }

class StatusTaskModule extends React.Component<StatusTaskModuleProps, IStatusState> {
    readonly localize: TFunction;
    private initMessage = {
        id: "",
        title: "",
        buttons: "[]",
        openFrom: ""
    };

    private card: any;

    constructor(props: StatusTaskModuleProps) {
        super(props);

        this.localize = this.props.t;

        this.state = {
            message: this.initMessage,
            loader: true,
            page: "ViewStatus",
            teamId: '',
            templateType: "",
            questionAnswer: [],
            id: ""
        };
    }

    public componentDidMount() {
        let params = this.props.match.params;
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.setState({
                teamId: context.teamId,
            });
        });

        if ('id' in params) {
            let id = params['id'];
            const parameter = new URLSearchParams(this.props.location.search);
            this.setState({
                openFrom: parameter.get('openfrom')
            }, () => {
                this.getItem(id).then(() => {
                    this.setState({
                        loader: false
                    }, () => {
                        this.card = (this.state.templateType === this.localize("ImageUpload")) ? getInitAdaptiveCard(this.localize) : (this.state.templateType === this.localize("PDFUpload")) ? getInitAdaptiveCardPDFUpload(this.localize) : (this.state.templateType === this.localize("Q&AUpload")) ? getInitAdaptiveCardQuestionAnswer(this.localize) : getInitAdaptiveCardEmailTemplate(this.localize);
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
                            setCardPartQuestionAnswer(this.card, this.state.questionAnswer, this.localize, this.state.message.title, this.state.message.author, teamsName); //update the adaptive cards
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
                                "url": this.state.message.buttonLink,
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
                        let link = this.state.message.buttonLink;
                        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
                    });
                });
            })

        }
    }

    private getItem = async (id: number) => {
        try {
            const response = (this.state.openFrom === "recall") ? await getRecallNotification(id) : await getSentNotification(id);
            console.log("response status", response)
            response.data.sendingDuration = formatDuration(response.data.sendingStartedDate, response.data.sentDate);
            response.data.sendingStartedDate = formatDate(response.data.sendingStartedDate);
            response.data.sentDate = formatDate(response.data.sentDate);
            response.data.succeeded = formatNumber(response.data.succeeded);
            response.data.failed = formatNumber(response.data.failed);
            console.log("view status", response.data)
            response.data.unknown = response.data.unknown && formatNumber(response.data.unknown);
            if (response.data.templateType === this.localize("Q&AUpload")) {
                if (response.data.summary !== "") {
                    this.setState({
                        questionAnswer: JSON.parse(response.data.summary)
                    });
                }
                this.setState({
                    adaptiveCardContent: JSON.parse(response.data.adaptiveCardContent),
                    templateType: response.data.templateType,
                }, () => {
                    this.setState({
                        notificationId: this.state.adaptiveCardContent.actions[0].data.NotificationId,
                        id: response.data.id,
                    }, () => {
                        this.exportData(this.state.id)
                        this.surveyEexportData(this.state.notificationId)
                    })

                })
            }
            else {
                this.setState({
                    id: response.data.id,
                    templateType: response.data.templateType
                }, () => {
                    this.exportData(this.state.id)
                })
            }
            this.setState({
                message: response.data,
                totalReadReceipt: response.data.totalReadReceipt
            });
        } catch (error) {
            return error;
        }
    }

    private exportData = async (id: any) => {

        const response = (this.state.openFrom === "recall") ? await recallDataExport(id) : await reactionexport(id);
        console.log("reaction export result", response);
        if (response.status === 200) {
            const downloadDataList = (this.state.openFrom !== "recall") ? response.data.map((e: any) => {
                let b = {
                    "Tenant": e.tenantName,
                    "Recipient Name": e.name,
                    "Recipient Email": e.email,
                    "Receipt Date": formatDate(e.sentDate),
                    "Reads On": e.readReceiptDate? formatDate(e.readReceiptDate) : "-",
                    "Reaction": e.reaction ? e.reaction : "-",
                    "Reacted On": e.reactionDate ? formatDate(e.reactionDate) : "-"

                }
                return b
            }): response.data.map((e: any) => {
                let b = {
                    "Tenant": e.tenantName,
                    "Recipient Name": e.name + " (" + e.email + ")",
                    "Receipt Date": formatDate(e.sentDate),
                    "Reads On":  e.readReceiptDate? formatDate(e.readReceiptDate) : "-",

                }
                return b
            })
            this.setState({
                downloadData: downloadDataList
            })
        }
    }

    private surveyEexportData = async (id: any) => {
        const response = await surveyexport(id);
        console.log("servey export result", response)
        if (response.data) {
            this.setState({
                surveyDownloadData: response.data
            })
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
            if (this.state.page === "ViewStatus") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half" className="formContentContainer">
                                    <Flex column>
                                        <div className="contentField">
                                            <h3>{this.localize("TitleText")}</h3>
                                            <span>{this.state.message.title}</span>
                                        </div>
                                        {(this.state.openFrom !== "recall") && <div className="contentField">
                                            <h3>{this.localize("ReadReceipt")}</h3>
                                            <span>{this.state.totalReadReceipt}</span>
                                        </div>}
                                        {(this.state.openFrom !== "recall") ? <div className="contentField">
                                            <h3>{this.localize("SendingStarted")}</h3>
                                            <span>{this.state.message.sendingStartedDate}</span>
                                        </div> : <div className="contentField">
                                            <h3>{this.localize("Recalled")}</h3>
                                            <span>{formatDate(this.state.message.recalledDate)}</span>
                                        </div>}
                                        {(this.state.openFrom !== "recall") && <div className="contentField">
                                            <h3>{this.localize("Completed")}</h3>
                                            <span>{this.state.message.sentDate}</span>
                                        </div>}
                                        {(this.state.openFrom !== "recall") && <div className="contentField">
                                            <h3>{this.localize("Duration")}</h3>
                                            <span>{this.state.message.sendingDuration}</span>
                                        </div>}
                                        <div className="contentField">
                                            <h3>{this.localize("Results")}</h3>
                                            <label>{this.localize("Success", { "SuccessCount": this.state.message.succeeded })}</label>
                                            <br />
                                            <label>{this.localize("Failure", { "FailureCount": this.state.message.failed })}</label>
                                            <br />
                                            {this.state.message.unknown &&
                                                <>
                                                    <label>{this.localize("Unknown", { "UnknownCount": this.state.message.unknown })}</label>
                                                </>
                                            }
                                        </div>
                                        <div className="contentField">
                                            <h3>{this.localize("Important")}</h3>
                                            <label>{this.renderImportant()}</label>
                                        </div>
                                        <div className="contentField">
                                            <h3>{this.localize("Tenant")}</h3>
                                            <label>{this.state.message.tenantName}</label>
                                        </div>
                                        <div className="contentField">
                                            {this.renderAudienceSelection()}
                                        </div>
                                       
                                        <div className="contentField">
                                            {this.renderErrorMessage()}
                                        </div>
                                        <div className="contentField">
                                            {this.renderWarningMessage()}
                                        </div>
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>
                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <div className={this.state.message.canDownload ? "" : "disabled"}>
                                    <Flex className="buttonContainer" gap="gap.small">
                                        <Flex.Item push>
                                            <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={this.localize("ExportLabel")} labelPosition="end" />
                                        </Flex.Item>
                                        <Flex.Item>
                                            <TooltipHost content={!this.state.message.sendingCompleted ? "" : (this.state.message.canDownload ? "" : this.localize("ExportButtonProgressText"))} calloutProps={{ gapSpace: 0 }}>
                                                {/* <Button icon={<DownloadIcon size="medium" />} disabled={this.state.message.canDownload || !this.state.message.sendingCompleted} content={this.localize("ExportButtonText")} id="exportBtn" onClick={this.onExport} primary /> */}
                                                <Dialog
                                                    closeOnOutsideClick={false}
                                                    confirmButton="ok"
                                                    content={this.localize("ExportMessage")}
                                                    trigger={<Button primary className="exportBtn" icon={<DownloadIcon size="medium" />} disabled={(this.state.downloadData && this.state.downloadData.length > 0) ? false : true} content={this.localize("ExportButtonText")}>
                                                        {this.state.downloadData && <CSVLink data={this.state.downloadData} filename={this.state.message.title + new Date().toDateString() + ".csv"}><DownloadIcon size="medium" styles={{ marginRight: "5px" }} />{this.localize("ExportButtonText")}</CSVLink>}
                                                    </Button>}
                                                />
                                                {/* <Button primary className="exportBtn" icon={<DownloadIcon size="medium" />} disabled={(this.state.downloadData && this.state.downloadData.length > 0) ? false : true} content={this.localize("ExportButtonText")}>
                                                    {this.state.downloadData && <CSVLink data={this.state.downloadData} filename={this.state.message.title+ new Date().toDateString() + ".csv"}><DownloadIcon size="medium" styles={{marginRight:"5px"}}/>{this.localize("ExportButtonText")}</CSVLink>}
                                                </Button> */}
                                                {(this.state.templateType === this.localize("Q&AUpload")) && (this.state.openFrom !== "recall") && 
                                                    <Dialog
                                                        closeOnOutsideClick={false}
                                                        confirmButton="ok"
                                                        content={this.localize("ExportMessage")}
                                                        trigger={<Button primary className="exportBtn" icon={<DownloadIcon size="medium" />} disabled={(this.state.surveyDownloadData && this.state.surveyDownloadData.length > 0) ? false : true} content={this.localize("SurveyExportButtonText")}>
                                                            {this.state.surveyDownloadData && <CSVLink data={this.state.surveyDownloadData} filename={this.state.message.title + new Date().toDateString() + ".csv"}><DownloadIcon size="medium" styles={{ marginRight: "5px" }} />{this.localize("SurveyExportButtonText")}</CSVLink>}
                                                        </Button>}
                                                    />}
                                            </TooltipHost>
                                        </Flex.Item>
                                    </Flex>
                                </div>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
            else if (this.state.page === "SuccessPage") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                            <div className="displayMessageField">
                                <br />
                                <br />
                                <div><span><AcceptIcon className="iconStyle" xSpacing="before" size="largest" outline /></span>
                                    <h1>{this.localize("ExportQueueTitle")}</h1></div>
                                <span>{this.localize("ExportQueueSuccessMessage1")}</span>
                                <br />
                                <br />
                                <span>{this.localize("ExportQueueSuccessMessage2")}</span>
                                <br />
                                <span>{this.localize("ExportQueueSuccessMessage3")}</span>
                            </div>
                            <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
                                <Flex className="buttonContainer">
                                    <Button content={this.localize("CloseText")} id="closeBtn" onClick={this.onClose} primary />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
            else {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                            <div className="displayMessageField">
                                <br />
                                <br />
                                <div><span></span>
                                    <h1 className="light">{this.localize("ExportErrorTitle")}</h1></div>
                                <span>{this.localize("ExportErrorMessage")}</span>
                            </div>
                            <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
                                <Flex className="buttonContainer">
                                    <Button content={this.localize("CloseText")} id="closeBtn" onClick={this.onClose} primary />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
        }
    }

    private onClose = () => {
        microsoftTeams.tasks.submitTask();
    }

    private onExport = async () => {
        let spanner = document.getElementsByClassName("sendingLoader");
        spanner[0].classList.remove("hiddenLoader");
        let payload = {
            id: this.state.message.id,
            teamId: this.state.teamId
        };
        await exportNotification(payload).then(() => {
            this.setState({ page: "SuccessPage" });
        }).catch(() => {
            this.setState({ page: "ErrorPage" });
        });
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
        if (this.state.message.teamNames && this.state.message.teamNames.length > 0) {
            return (
                <div>
                    <h3>{this.localize("SentToGeneralChannel")}</h3>
                    <List items={this.getItemList(this.state.message.teamNames)} />
                </div>);
        } else if (this.state.message.rosterNames && this.state.message.rosterNames.length > 0) {
            return (
                <div>
                    <h3>{this.localize("SentToRosters")}</h3>
                    <List items={this.getItemList(this.state.message.rosterNames)} />
                </div>);
        } else if (this.state.message.groupNames && this.state.message.groupNames.length > 0) {
            return (
                <div>
                    <h3>{this.localize("SentToGroups1")}</h3>
                    <span>{this.localize("SentToGroups2")}</span>
                    <List items={this.getItemList(this.state.message.groupNames)} />
                </div>);
        } else if (this.state.message.allUsers) {
            return (
                <div>
                    <h3>{this.localize("SendToAllUsers")}</h3>
                </div>);
        } else {
            return (<div></div>);
        }
    }
    private renderErrorMessage = () => {
        if (this.state.message.errorMessage) {
            return (
                <div>
                    <h3>{this.localize("Errors")}</h3>
                    <span>{this.state.message.errorMessage}</span>
                </div>
            );
        } else {
            return (<div></div>);
        }
    }

    private renderWarningMessage = () => {
        if (this.state.message.warningMessage) {
            return (
                <div>
                    <h3>{this.localize("Warnings")}</h3>
                    <span>{this.state.message.warningMessage}</span>
                </div>
            );
        } else {
            return (<div></div>);
        }
    }
}

const StatusTaskModuleWithTranslation = withTranslation()(StatusTaskModule);
export default StatusTaskModuleWithTranslation;