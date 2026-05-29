// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';

import { Button, Loader, Dropdown, Text, Flex, Input, TextArea, RadioGroup, Checkbox, Datepicker, FormDropdown, Popup } from '@fluentui/react-northstar'
import { TrashCanIcon, AddIcon, FilesUploadIcon } from '@fluentui/react-icons-northstar'

import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import { TFunction } from "i18next";
import * as microsoftTeams from "@microsoft/teams-js";

import * as AdaptiveCards from "adaptivecards";

import { getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary, setCardAuthor, setCardBtns, setCardTeamsName } from '../AdaptiveCard/adaptiveCard';
import { getInitAdaptiveCardPDFUpload, setCardTitlePDFUpload, setCardImageLinkPDFUpload, setCardPdfNamePDFUpload, setCardSummaryPDFUpload, setCardAuthorPDFUpload, setCardBtnsPDFUpload, setCardTeamsNamePDFUpload } from '../AdaptiveCard/adaptiveCardPDFUpload';
import { getInitAdaptiveCardQuestionAnswer, setCardTitleQuestionAnswer, setCardAuthorQuestionAnswer, setCardPartQuestionAnswer, setCardTeamsNameQuestionAnswere } from '../AdaptiveCard/adaptiveCardQuestionAnswer';
import { getInitAdaptiveCardEmailTemplate, setCardEmailTemplate } from '../AdaptiveCard/adaptiveCardEmailTemplate';

import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, searchGroups, getGroups, verifyGroupAccess, sendPdfFile, getSisterTenant, getAllTenant } from '../../apis/messageListApi';

import { getBaseUrl } from '../../configVariables';
import { ImageUtil } from '../../utility/imageutility';

import './newMessage.scss';
import './teamTheme.scss';


//hours to be chosen when scheduling messages
const hours = ["00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
    "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23",
];

//minutes to be chosen when scheduling messages
const minutes = ["00", "05", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55",
];

//coeficient to round dates to the next 5 minutes
const coeff = 1000 * 60 * 5;

const pdfImgUrl = getBaseUrl() + "/image/pdfImage.png";

type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: {
        id: string,
    },
}

export interface IDraftMessage {
    id?: string,
    title: string,
    imageLink?: any,
    summary?: string,
    author: string,
    buttonTitle?: string,
    buttonLink?: string,
    teams: any[],
    rosters: any[],
    groups: any[],
    allUsers: boolean,
    isImportant: boolean, // indicates if the message is important
    isScheduled: boolean, // indicates if the message is scheduled
    ScheduledDate: Date, // stores the scheduled date
    Buttons: string, // stores tha card buttons (JSON)
    TemplateType: string,
    SendTypeId?: any,
    TenantId?: any,
    EmailBody?: any,
    EmailTitle: string,
    AdaptiveCardContent?: string,
    CsvLink?: any,
    TenantName?: string,
    AuthorTeamId?: any,
    AuthorTeamName?: string,
    AuthorChannelId?: any,
    AuthorChannelName?: string,
}

export interface formState {
    title: string,
    summary?: string,
    btnLink?: string,
    imageLink?: any,
    btnTitle?: string,
    author: string,
    card?: any,
    page: string,
    teamsOptionSelected: boolean,
    rostersOptionSelected: boolean,
    allUsersOptionSelected: boolean,
    sisterTenantOptionSelected: boolean,
    csvLink?: any,
    csvOptionSelected: boolean,
    groupsOptionSelected: boolean,
    teams?: any[],
    groups?: any[],
    exists?: boolean,
    messageId?: any,
    loader: boolean,
    groupAccess: boolean,
    loading: boolean,
    noResultMessage: string,
    unstablePinned?: boolean,
    selectedTeamsNum: number,
    selectedRostersNum: number,
    selectedGroupsNum: number,
    selectedRadioBtn?: string,
    selectedTeams: dropdownItem[],
    selectedRosters: dropdownItem[],
    selectedGroups: dropdownItem[],
    errorImageUrlMessage: string,
    errorButtonUrlMessage: string,
    selectedSchedule: boolean, //status of the scheduler checkbox
    selectedImportant: boolean, //status of the importance selection on the interface
    scheduledDate: string, //stores the scheduled date in string format
    DMY: Date, //scheduled date in date format
    DMYHour: string, //hour selected
    DMYMins: string, //mins selected
    futuredate: boolean, //if the date is in the future (valid schedule)
    values: any[], //button values collection
    templateType?: any,
    questionTypeDropDownInput?: any,
    questionTypeSelectedValue?: any,
    questionTypeSelectedValueDisable: boolean,
    selectedFileName?: any,
    questionAnswer: any[],
    pdfFile?: any,
    addQuestionError?: any,
    backButtonShow?: any,
    tenantId?: any,
    emailBodyContent?: any,
    emailTitleText?: any,
    emailFileTitle?: any,
    sistertenantId?: any,
    imageHeight?: any,
    imageWidth?: any,
    allTenantList?: any,
    fileUploadText?: string,
    selectedTenantName?: any,
    AuthorTeamId?: any,
    AuthorTeamName?: string,
    AuthorChannelId?: any,
    AuthorChannelName?: string,
    errorEmailUrlMessage?: string,
}

export interface INewMessageProps extends RouteComponentProps, WithTranslation {
    getDraftMessagesList?: any;
}

class NewMessage extends React.Component<INewMessageProps, formState> {
    readonly localize: TFunction;
    private card: any;
    fileInput: any;

    constructor(props: INewMessageProps) {
        super(props);
        this.localize = this.props.t;

        var TempDate = this.getRoundedDate(5, this.getDateObject()); //get the current date
        this.state = {
            title: "",
            summary: "",
            author: "",
            btnLink: "",
            imageLink: "",
            selectedFileName: "",
            btnTitle: "",
            pdfFile: "",
            page: "CardCreation",
            teamsOptionSelected: true,
            rostersOptionSelected: false,
            allUsersOptionSelected: false,
            sisterTenantOptionSelected: false,
            csvOptionSelected: false,
            csvLink: "",
            groupsOptionSelected: false,
            messageId: "",
            loader: true,
            groupAccess: false,
            loading: false,
            noResultMessage: "",
            unstablePinned: true,
            selectedTeamsNum: 0,
            selectedRostersNum: 0,
            selectedGroupsNum: 0,
            selectedTeams: [],
            selectedRosters: [],
            selectedGroups: [],
            errorImageUrlMessage: "",
            errorButtonUrlMessage: "",
            selectedSchedule: false, //scheduler option is disabled by default
            selectedImportant: false, //important flag for the msg is false by default
            scheduledDate: TempDate.toUTCString(), //current date in UTC string format
            DMY: TempDate, //current date in Date format
            DMYHour: this.getDateHour(TempDate.toUTCString()), //initialize with the current hour (rounded up)
            DMYMins: this.getDateMins(TempDate.toUTCString()), //initialize with the current minute (rounded up)
            futuredate: false, //by default the date is not in the future
            values: [], //by default there are no buttons on the adaptive card.
            questionAnswer: [],
            questionTypeDropDownInput: [this.localize("DescriptiveQuestion"), this.localize("MCQ")], //for survey template dropdown input
            addQuestionError: "",
            backButtonShow: false,
            tenantId: "",
            emailTitleText: false,
            emailFileTitle: "",
            questionTypeSelectedValueDisable: false,
            sistertenantId: "",
            allTenantList: [],
            fileUploadText: "",
            selectedTenantName: "",
            AuthorTeamId: "",
            AuthorTeamName: "",
            AuthorChannelId: "",
            AuthorChannelName: "",
            errorEmailUrlMessage: "",
            emailBodyContent:""
        }
        this.fileInput = React.createRef();
        this.handleImagePDFSelection = this.handleImagePDFSelection.bind(this);
    }


    public async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            console.log("Context", context)
            this.setState({
                AuthorTeamId: context.teamId,
                AuthorTeamName: context.teamName,
                AuthorChannelId: context.channelId,
                AuthorChannelName: context.channelName,
            }, () => {
                //- Handle the Esc key
                document.addEventListener("keydown", this.escFunction, false);
                let params = this.props.match.params;
                this.setGroupAccess();
                this.getSisterTenantId();




                if ('id' in params) {
                    let id = params['id'];
                    this.getItem(id).then(() => {

                        this.setState({
                            exists: true,
                            messageId: id,
                            selectedSchedule: this.state.selectedSchedule,
                            selectedImportant: this.state.selectedImportant,
                            scheduledDate: this.state.scheduledDate,
                            DMY: this.getDateObject(this.state.scheduledDate),
                            DMYHour: this.getDateHour(this.state.scheduledDate),
                            DMYMins: this.getDateMins(this.state.scheduledDate),
                            values: this.state.values,
                            questionTypeSelectedValue: this.state.questionTypeDropDownInput[0],
                        })
                    });
                    this.getAllTenant("Edit");
                }

                else {
                    if (this.props.location.state) {
                        this.setState({
                            templateType: this.props.location.state.data,
                            questionTypeSelectedValue: this.state.questionTypeDropDownInput[0],
                            backButtonShow: true,
                            exists: false,
                            loader: false,
                        }, () => {
                            console.log("sayan", this.state.templateType)
                            this.card = (this.state.templateType === this.localize("ImageUpload")) ? getInitAdaptiveCard(this.localize) : (this.state.templateType === this.localize("PDFUpload")) ? getInitAdaptiveCardPDFUpload(this.localize) : (this.state.templateType === this.localize("Q&AUpload")) ? getInitAdaptiveCardQuestionAnswer(this.localize) : getInitAdaptiveCardEmailTemplate(this.localize);
                            this.setState({
                                card: this.card,
                                exists: false,
                                loader: false,
                                selectedRadioBtn: ((this.state.templateType === this.localize("ImageUpload")) || (this.state.templateType === this.localize("PDFUpload"))) ? "teams" : "rosters",
                                teamsOptionSelected: ((this.state.templateType === this.localize("ImageUpload")) || (this.state.templateType === this.localize("PDFUpload"))) ? true : false,
                                rostersOptionSelected: ((this.state.templateType === this.localize("ImageUpload")) || (this.state.templateType === this.localize("PDFUpload"))) ? false : true,
                                allUsersOptionSelected: false,
                                sisterTenantOptionSelected: false,
                                csvOptionSelected: false,
                                groupsOptionSelected: false
                            }, () => {
                                this.setDefaultCard(this.card);
                                let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                                adaptiveCard.parse(this.state.card);
                                let renderedCard = adaptiveCard.render();
                                document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                                if (this.state.btnLink) {
                                    let link = this.state.btnLink;
                                    adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                                }
                            })
                        })
                        this.getAllTenant("Create");
                    }
                }

            })
        });
    }


    fileUpload() {
        (document.getElementById('upload') as HTMLInputElement).click()
    };

    onFileChoose(event: any) {
        console.log(event.target.files[0])
        var csvFormData = new FormData()
        csvFormData.append("file", event.target.files[0]);
        this.uploadCSVfile(csvFormData)

    }

    private uploadCSVfile = async (file: any) => {
        const response = await sendPdfFile(file);
        if (response.data) {
            console.log("link res", response.data)
            this.setState({
                csvLink: response.data,
            })
        }
    }

    //function to handle the selection of the OS file upload box
    private handleImagePDFSelection() {
        //get the first file selected
        const file = this.fileInput.current.files[0];
        if (file) { //if we have a file
            if (this.state.templateType === this.localize("ImageUpload")) {  //for image upload
                this.imageCheck(file)
            }
            else if (this.state.templateType === this.localize("PDFUpload")) {   //for pdf upload
                if (file.size < 15728640) {
                    var pdfFormData = new FormData()
                    pdfFormData.append("file", file);
                    this.getUploadedFileURL(pdfFormData, file)
                }
                else {
                    this.setState({
                        errorImageUrlMessage: this.localize("ErrorPDFTooBig")
                    });
                }

            }
            else if (this.state.templateType === this.localize("EmailUpload")) {
                this.setState({
                    emailTitleText: false,
                    fileUploadText: ""
                })
                console.log("email", file)
                if (file.size < 15728640) {

                    var emailFormData = new FormData()
                    emailFormData.append("file", file);
                    this.getUploadedFileURL(emailFormData, file)

                }
                else {
                    this.setState({
                        errorEmailUrlMessage: this.localize("ErrorEmailTemplateTooBig")
                    });
                }

            }
        }

    }

    onImageFileChooseEmail(event: any) {
        // console.log(event.target.files[0])
        this.imageCheck(event.target.files[0])
    }

    private imageCheck = async (file: any) => {
        var reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
            if (reader.result) {
                const image = new Image();
                image.src = reader.result.toString()
                image.onload = () => {
                    if (image.height <= 1024 && image.width <= 1024 && file.size < 1048576) {
                        var imageFormData = new FormData()
                        imageFormData.append("file", file);
                        this.getUploadedFileURL(imageFormData, file)
                    }
                    else {
                        this.setState({
                            errorImageUrlMessage: this.localize("ErrorImageTooBig")
                        });
                    }
                };
            }
        };
        reader.onerror = function (error) {
            console.log('Error: ', error);
        };
    }


    private getUploadedFileURL = async (formdata: any, file: any) => {
        const response = await sendPdfFile(formdata);
        if (response.data) {
            // console.log("link res", response.data);
            if (this.state.templateType === this.localize("PDFUpload")) {
                this.setState({
                    imageLink: response.data
                })
                setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                let pdfLink = "[View PDF](" + response.data + ")"
                setCardPdfNamePDFUpload(this.card, pdfLink)
                this.updateCard();
            }
            else if (this.state.templateType === this.localize("EmailUpload")) {
                if (file.type.includes('image')) {
                    this.setState({
                        emailBodyContent: response.data
                    })
                    let EmailLink = "[" + this.state.emailFileTitle + "](" + this.state.imageLink + ")"
                    let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
                    setCardEmailTemplate(this.card, teamsName, response.data, this.state.title, this.state.author, EmailLink, this.state.summary)
                    this.updateCard();
                }
                else {
                    this.setState({
                        imageLink: response.data,
                        emailTitleText: true,
                        fileUploadText: this.localize("UploadCompleteText")
                    })
                    let EmailLink = "[" + this.state.emailFileTitle + "](" + response.data + ")"
                    let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
                    setCardEmailTemplate(this.card, teamsName, this.state.emailBodyContent, this.state.title, this.state.author, EmailLink, this.state.summary)
                    this.updateCard();
                }

            }
            else if (this.state.templateType === this.localize("ImageUpload")) {
                this.setState({
                    imageLink: response.data
                })
                setCardImageLink(this.card, response.data);
                this.updateCard();
            }


        }

    }




    //Function calling a click event on a hidden file input
    private handleUploadClick = (event: any) => {
        //reset the error message and the image link as the upload will reset them potentially
        this.setState({
            errorImageUrlMessage: "",
            imageLink: ""
        });
        //fire the fileinput click event and run the handleImagePDFSelection function
        this.fileInput.current.click();
    };

    private getSisterTenantId = async () => {
        const response = await getSisterTenant();
        this.setState({
            sistertenantId: response.data
        })
    }

    private getAllTenant = async (type: any) => {
        const response = await getAllTenant();
        if (response.data) {
            let result = response.data.map((a: any) => {
                let b = {
                    header: a.name,
                    tenantId: a.rowKey
                }
                return b
            })
            this.setState({
                allTenantList: result
            })
            if (type === "Create") {

                let primaryTenant = response.data.filter((e: any) => e.isPrimary === true).map((a: any) => {
                    let b = {
                        header: a.name,
                        tenantId: a.rowKey
                    }
                    return b
                })
                this.selectTenant(primaryTenant[0])
            }
        }
    }

    selectTenant(data: any) {
        console.log("Selected Tenant", data)
        this.setState({
            selectedTenantName: data.header,
            tenantId: data.tenantId
        }, () => {
            console.log("Selected Tenant", this.state.tenantId, this.state.selectedTenantName)
            this.getTeamList()
        })
    }


    private getTeamList = async () => {
        try {
            const response = await getTeams(this.state.tenantId);
            console.log("Team List", response)
            this.setState({
                teams: response.data
            });
        } catch (error) {
            return error;
        }
    }

    private makeDropdownItems = (items: any[] | undefined) => {
        const resultedTeams: dropdownItem[] = [];
        if (items) {
            items.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id
                    },

                });
            });
        }
        return resultedTeams;
    }

    private makeDropdownItemList = (items: any[], fromItems: any[] | undefined) => {
        const dropdownItemList: dropdownItem[] = [];
        items.forEach(element =>
            dropdownItemList.push(
                typeof element !== "string" ? element : {
                    key: fromItems!.find(x => x.id === element).id,
                    header: fromItems!.find(x => x.id === element).name,
                    image: ImageUtil.makeInitialImage(fromItems!.find(x => x.id === element).name),
                    team: {
                        id: element
                    }
                })
        );
        return dropdownItemList;
    }

    public setDefaultCard = (card: any) => {

        const titleAsString = this.localize("TitleText");
        const summaryAsString = this.localize("Summary");
        const authorAsString = this.localize("Author1");
        const buttonTitleAsString = this.localize("ButtonTitle");
        const email = this.localize("Email");
        const imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
        const teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName

        if (this.state.templateType === this.localize("ImageUpload")) {
            setCardTeamsName(card, teamsName)
            setCardTitle(card, titleAsString);
            // let imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
            setCardImageLink(card, imgUrl);
            setCardSummary(card, summaryAsString);
            setCardAuthor(card, authorAsString);
            setCardBtns(card, [{
                "type": "Action.OpenUrl",
                "title": "Button",
                "url": ""
            }]);
        }
        else if (this.state.templateType === this.localize("PDFUpload")) {
            setCardTeamsNamePDFUpload(card, teamsName)
            setCardTitlePDFUpload(card, titleAsString);
            setCardSummaryPDFUpload(card, summaryAsString);
            setCardAuthorPDFUpload(card, authorAsString);
            setCardBtnsPDFUpload(card, [{
                "type": "Action.OpenUrl",
                "title": "Button",
                "url": ""
            }]);

        }
        else if (this.state.templateType === this.localize("Q&AUpload")) {
            setCardTeamsNameQuestionAnswere(card, teamsName)
            setCardTitleQuestionAnswer(card, titleAsString);
            setCardAuthorQuestionAnswer(card, authorAsString);
        }

        else {
            setCardEmailTemplate(card, teamsName, imgUrl, titleAsString, authorAsString, email, summaryAsString)
        }

    }



    private getGroupItems() {
        if (this.state.groups) {
            return this.makeDropdownItems(this.state.groups);
        }
        const dropdownItems: dropdownItem[] = [];
        return dropdownItems;
    }

    private setGroupAccess = async () => {
        await verifyGroupAccess().then(() => {
            this.setState({
                groupAccess: true
            });
        }).catch((error) => {
            const errorStatus = error.response.status;
            if (errorStatus === 403) {
                this.setState({
                    groupAccess: false
                });
            }
            else {
                throw error;
            }
        });
    }

    private getGroupData = async (id: number) => {
        try {
            const response = await getGroups(id, this.state.tenantId);
            console.log("get group data response", response)
            this.setState({
                groups: response.data
            });
        }
        catch (error) {
            return error;
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getDraftNotification(id);
            console.log("getItem response", response)
            const draftMessageDetail = response.data;
            let selectedRadioButton = "teams";
            if (draftMessageDetail.rosters.length > 0) {
                selectedRadioButton = "rosters";
            }
            else if (draftMessageDetail.groups.length > 0) {
                selectedRadioButton = "groups";
            }
            else if (draftMessageDetail.allUsers) {
                if (draftMessageDetail.sendTypeId === '3') {
                    selectedRadioButton = "allUsers";
                }
                else {
                    selectedRadioButton = "sistertenant";
                }

            }
            else if (draftMessageDetail.csvLink) {
                selectedRadioButton = "csv";
            }

            // set state based on values returned 
            this.setState({
                teamsOptionSelected: draftMessageDetail.teams.length > 0,
                selectedTeamsNum: draftMessageDetail.teams.length,
                rostersOptionSelected: draftMessageDetail.rosters.length > 0,
                selectedRostersNum: draftMessageDetail.rosters.length,
                groupsOptionSelected: draftMessageDetail.groups.length > 0,
                selectedGroupsNum: draftMessageDetail.groups.length,
                selectedRadioBtn: selectedRadioButton,
                selectedTeams: draftMessageDetail.teams,
                selectedRosters: draftMessageDetail.rosters,
                selectedGroups: draftMessageDetail.groups,
                selectedSchedule: draftMessageDetail.isScheduled,
                selectedImportant: draftMessageDetail.isImportant,
                scheduledDate: draftMessageDetail.scheduledDate,
                templateType: draftMessageDetail.templateType,
                tenantId: draftMessageDetail.tenantId,
                selectedTenantName: draftMessageDetail.tenantName,
                AuthorTeamId: draftMessageDetail.authorTeamId,
                AuthorTeamName: draftMessageDetail.authorTeamName,
                AuthorChannelId: draftMessageDetail.authorChannelId,
                AuthorChannelName: draftMessageDetail.authorChannelName,
            }, () => {

                this.getTeamList().then(() => {
                    const selectedTeams = this.makeDropdownItemList(this.state.selectedTeams, this.state.teams);
                    const selectedRosters = this.makeDropdownItemList(this.state.selectedRosters, this.state.teams);
                    this.setState({
                        selectedTeams: selectedTeams,
                        selectedRosters: selectedRosters,
                    })
                })

                this.getGroupData(id).then(() => {
                    const selectedGroups = this.makeDropdownItems(this.state.groups);
                    this.setState({
                        selectedGroups: selectedGroups
                    })
                });

                this.card = (this.state.templateType === this.localize("ImageUpload")) ? getInitAdaptiveCard(this.localize) : (this.state.templateType === this.localize("PDFUpload")) ? getInitAdaptiveCardPDFUpload(this.localize) : (this.state.templateType === this.localize("Q&AUpload")) ? getInitAdaptiveCardQuestionAnswer(this.localize) : getInitAdaptiveCardEmailTemplate(this.localize);
                // set card properties
                if (this.state.templateType === this.localize("ImageUpload")) {
                    let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
                    setCardTeamsName(this.card, teamsName);
                    setCardTitle(this.card, draftMessageDetail.title);
                    setCardImageLink(this.card, draftMessageDetail.imageLink);
                    setCardSummary(this.card, draftMessageDetail.summary);
                    setCardAuthor(this.card, draftMessageDetail.author);
                }
                else if (this.state.templateType === this.localize("PDFUpload")) {
                    let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
                    setCardTeamsNamePDFUpload(this.card, teamsName);
                    setCardTitlePDFUpload(this.card, draftMessageDetail.title);
                    setCardSummaryPDFUpload(this.card, draftMessageDetail.summary);
                    setCardAuthorPDFUpload(this.card, draftMessageDetail.author);
                    if (draftMessageDetail.imageLink !== "") {
                        setCardImageLinkPDFUpload(this.card, pdfImgUrl);

                        let pdfLink = "[View PDF](" + draftMessageDetail.imageLink + ")"
                        setCardPdfNamePDFUpload(this.card, pdfLink)

                    }

                }
                else if (this.state.templateType === this.localize("Q&AUpload")) {
                    let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
                    setCardTeamsNameQuestionAnswere(this.card, teamsName);
                    setCardTitleQuestionAnswer(this.card, draftMessageDetail.title);
                    setCardAuthorQuestionAnswer(this.card, draftMessageDetail.author);
                    if (draftMessageDetail.summary !== "") {

                        this.setState({
                            questionAnswer: JSON.parse(draftMessageDetail.summary)
                        }, () => {
                            setCardPartQuestionAnswer(this.card, this.state.questionAnswer, this.localize, draftMessageDetail.title, draftMessageDetail.author, teamsName); //update the adaptive card
                        });
                    }
                }
                else {

                    if (draftMessageDetail.imageLink !== "") {
                        this.setState({
                            emailTitleText: true
                        })
                        let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
                        let EmailLink = "[" + draftMessageDetail.emailTitle + "](" + draftMessageDetail.imageLink + ")"
                        setCardEmailTemplate(this.card, teamsName, draftMessageDetail.emailBody, draftMessageDetail.title, draftMessageDetail.author, EmailLink, draftMessageDetail.summary)
                    }
                }

                // this is to ensure compatibility with older versions
                // if we get empty buttonsJSON and values on buttonTitle and buttonLink, we insert those to values
                // if not we just use values cause the JSON will be complete over there
                // console.log("getItem response button", draftMessageDetail.buttons)

                if (draftMessageDetail.buttonTitle && draftMessageDetail.buttonLink && !draftMessageDetail.buttons) {
                    this.setState({
                        values: [{
                            "type": "Action.OpenUrl",
                            "title": draftMessageDetail.buttonTitle,
                            "url": draftMessageDetail.buttonLink
                        }]
                    });
                }
                else {
                    // set the values state with the parse of the JSON recovered from the database
                    if (draftMessageDetail.buttons !== null) { //if the database value is not null, parse the JSON to create the button objects
                        this.setState({
                            values: JSON.parse(draftMessageDetail.buttons)
                        }, () => {
                            // set the card buttons collection based on the values collection
                            if (this.state.templateType === this.localize("ImageUpload")) {
                                setCardBtns(this.card, this.state.values);
                            }
                            else if (this.state.templateType === this.localize("PDFUpload")) {
                                setCardBtnsPDFUpload(this.card, this.state.values);
                            }
                        });
                    } else { //if the string is null, then initialize the empty collection 
                        this.setState({
                            values: []
                        });
                    }
                }
                this.setState({
                    title: draftMessageDetail.title,
                    summary: draftMessageDetail.summary,
                    btnLink: draftMessageDetail.buttonLink,
                    imageLink: draftMessageDetail.imageLink,
                    btnTitle: draftMessageDetail.buttonTitle,
                    author: draftMessageDetail.author,
                    allUsersOptionSelected: draftMessageDetail.allUsers,
                    sisterTenantOptionSelected: draftMessageDetail.sendTypeId === '5',
                    csvOptionSelected: draftMessageDetail.sendTypeId === '6',
                    csvLink: draftMessageDetail.sendTypeId === '6' ? draftMessageDetail.csvLink : "",
                    loader: false,
                    card: this.card,
                    emailFileTitle: draftMessageDetail.emailTitle,
                    emailBodyContent: draftMessageDetail.emailBody
                }, () => {
                    this.updateCard();
                });


            });
        } catch (error) {
            return error;
        }
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    private questionTypeHandleOnchange = (data: any) => {
        this.setState({
            questionTypeSelectedValue: data
        })
    }

    public render(): JSX.Element {
        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            if (this.state.page === "CardCreation") {
                return (
                    <div>
                        {(this.state.templateType === this.localize("ImageUpload") || this.state.templateType === this.localize("PDFUpload")) ? <div className="taskModule">
                            <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                                <Flex className="scrollableContent">
                                    <Flex.Item size="size.half">
                                        <Flex column className="formContentContainer">
                                            <Input className="inputField"
                                                value={this.state.title}
                                                label={this.localize("TitleText")}
                                                placeholder={this.localize("PlaceHolderTitle")}
                                                onChange={this.onTitleChanged}
                                                autoComplete="off"
                                                fluid
                                            />
                                            <Flex gap="gap.smaller" vAlign="end" className="inputField">

                                                {(this.state.templateType === this.localize("ImageUpload")) ? <Input
                                                    value={this.state.imageLink}
                                                    label={this.localize("ImageURL")}
                                                    placeholder={this.localize("ImageURLPlaceHolder")}
                                                    onChange={this.onImageLinkChanged}
                                                    error={!(this.state.errorImageUrlMessage === "")}
                                                    autoComplete="off"
                                                    fluid
                                                /> :
                                                    <Text size="medium" content={this.localize("PDFUploadText")} />
                                                }
                                                {(this.state.templateType === this.localize("ImageUpload")) ? <input type="file" accept=".png, .jpg, .jpeg, .gif"
                                                    style={{ display: 'none' }}
                                                    onChange={this.handleImagePDFSelection}
                                                    ref={this.fileInput} /> :
                                                    <input type="file" accept=".pdf"
                                                        style={{ display: 'none' }}
                                                        onChange={this.handleImagePDFSelection}
                                                        ref={this.fileInput} />
                                                }
                                                <Flex.Item push>
                                                    <Button circular onClick={this.handleUploadClick}
                                                        size="small"
                                                        icon={<FilesUploadIcon />}
                                                        title={this.localize("UploadText")}
                                                        styles={{ marginTop: "10px" }}
                                                    />
                                                </Flex.Item>
                                            </Flex>
                                            <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />

                                            <div className="textArea">
                                                <Text content={this.localize("Summary")} />
                                                <TextArea
                                                    autoFocus
                                                    placeholder={this.localize("Summary")}
                                                    value={this.state.summary}
                                                    onChange={this.onSummaryChanged}
                                                    fluid />
                                            </div>

                                            <Input className="inputField"
                                                value={this.state.author}
                                                label={this.localize("Author")}
                                                placeholder={this.localize("Author")}
                                                onChange={this.onAuthorChanged}
                                                autoComplete="off"
                                                fluid
                                            />
                                            <div className="textArea">
                                                <Flex gap="gap.large" vAlign="end">
                                                    <Text size="small" align="start" content={this.localize("Buttons")} />
                                                    <Flex.Item push >
                                                        <Button circular size="small" disabled={(this.state.values.length == 4) || !(this.state.errorButtonUrlMessage === "")} icon={<AddIcon />} title={this.localize("Add")} onClick={this.addClick.bind(this)} />
                                                    </Flex.Item>
                                                </Flex>
                                            </div>

                                            {this.createUI()}

                                            <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                        </Flex>
                                    </Flex.Item>
                                    <Flex.Item size="size.half">
                                        <div className="adaptiveCardContainer">
                                        </div>
                                    </Flex.Item>
                                </Flex>

                                <Flex className="footerContainer" vAlign="end" hAlign="end">
                                    <Button content={this.localize("Back")} onClick={() => this.props.history.push({ pathname: "/customNewmessage", state: { data: this.state.templateType } })} secondary />
                                    <Flex className="buttonContainer">
                                        <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                    </Flex>
                                </Flex>

                            </Flex>
                        </div> :
                            (this.state.templateType === this.localize("Q&AUpload")) ? <div className="taskModule">
                                <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                                    <Flex className="scrollableContent">
                                        <Flex.Item size="size.half">
                                            <Flex column className="formContentContainer">
                                                <Input className="inputField"
                                                    value={this.state.title}
                                                    label={this.localize("TitleText")}
                                                    placeholder={this.localize("PlaceHolderTitle")}
                                                    onChange={this.onTitleChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />
                                                <Input className="inputField"
                                                    value={this.state.author}
                                                    label={this.localize("Author")}
                                                    placeholder={this.localize("Author")}
                                                    onChange={this.onAuthorChanged}
                                                    autoComplete="off"
                                                    fluid
                                                />

                                                <FormDropdown className="inputField dropdown_style"
                                                    items={this.state.questionTypeDropDownInput}
                                                    label={this.localize("TypeofQuestionText")}
                                                    value={this.state.questionTypeSelectedValue}
                                                    placeholder={this.state.questionTypeDropDownInput[0]}

                                                    onChange={(event, { value }) => this.questionTypeHandleOnchange(value)}
                                                    fluid
                                                />

                                                <div className="textArea">
                                                    <Flex gap="gap.large" vAlign="end" styles={{ marginTop: "10px" }}>
                                                        <Flex.Item push >
                                                            <Button circular size="small" disabled={(this.state.questionAnswer.length == 8) || !(this.state.addQuestionError === "")} icon={<AddIcon />} title={this.localize("Add")} onClick={this.addQuestionAnswer.bind(this)} />
                                                        </Flex.Item>
                                                    </Flex>
                                                </div>

                                                {this.createQuestionAnswerUI()}

                                                <Text className={(this.state.addQuestionError === "") ? "hide" : "show"} error size="small" content={this.state.addQuestionError} />


                                            </Flex>
                                        </Flex.Item>
                                        <Flex.Item size="size.half">
                                            <div className="adaptiveCardContainer">
                                            </div>
                                        </Flex.Item>
                                    </Flex>

                                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                                        {this.state.backButtonShow && <Button content={this.localize("Back")} onClick={() => this.props.history.push({ pathname: "/customNewmessage", state: { data: this.state.templateType } })} secondary />}
                                        <Flex className="buttonContainer">
                                            <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                        </Flex>
                                    </Flex>

                                </Flex>
                            </div>
                                :
                                <div className="taskModule">
                                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                                        <Flex className="scrollableContent">
                                            <Flex.Item size="size.half">
                                                <Flex column className="formContentContainer">
                                                    <Input className="inputField"
                                                        value={this.state.title}
                                                        label={this.localize("TitleText")}
                                                        placeholder={this.localize("PlaceHolderTitle")}
                                                        onChange={this.onTitleChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField" styles={{ marginTop: "10px" }}>

                                                        <Text size="medium" content={this.localize("EmailTemplateImageUploadText")} />
                                                        <input type="file" id="upload" style={{ display: 'none' }} onChange={value => this.onImageFileChooseEmail(value)} accept=".png, .jpg, .jpeg, .gif" />
                                                        <Flex.Item push>
                                                            <Button circular onClick={() => this.fileUpload()}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title={this.localize("UploadText")}

                                                            />
                                                        </Flex.Item>
                                                    </Flex>
                                                    <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />
                                                    <Flex gap="gap.smaller" vAlign="end" className="inputField" styles={{ marginTop: "10px" }}>

                                                        <Text size="medium" content={this.localize("EmailUploadText")} />
                                                        <input type="file" accept=".html"
                                                            style={{ display: 'none' }}
                                                            onChange={this.handleImagePDFSelection}
                                                            ref={this.fileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={this.handleUploadClick}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title={this.localize("UploadText")}

                                                            />
                                                        </Flex.Item>
                                                    </Flex>
                                                    <Text className={(this.state.fileUploadText === "") ? "hide" : "show"} error size="small" content={this.state.fileUploadText} />
                                                    <Text className={(this.state.errorEmailUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorEmailUrlMessage} />
                                                    {this.state.emailTitleText && <Input className="inputField"
                                                        value={this.state.emailFileTitle}
                                                        label={this.localize("EmailTitle")}
                                                        placeholder={this.localize("EmailTitle")}
                                                        onChange={this.EmailTitle}
                                                        autoComplete="off"
                                                        fluid
                                                    />}
                                                    <div className="textArea">
                                                        <Text content={this.localize("Summary")} />
                                                        <TextArea
                                                            autoFocus
                                                            placeholder={this.localize("Summary")}
                                                            value={this.state.summary}
                                                            onChange={this.onSummaryChanged}
                                                            fluid />
                                                    </div>
                                                    <Input className="inputField"
                                                        value={this.state.author}
                                                        label={this.localize("Author")}
                                                        placeholder={this.localize("Author")}
                                                        onChange={this.onAuthorChanged}
                                                        autoComplete="off"
                                                        fluid
                                                    />

                                                </Flex>
                                            </Flex.Item>
                                            <Flex.Item size="size.half">
                                                <div className="adaptiveCardContainer">
                                                </div>
                                            </Flex.Item>
                                        </Flex>

                                        <Flex className="footerContainer" vAlign="end" hAlign="end">
                                            {this.state.backButtonShow && <Button content={this.localize("Back")} onClick={() => this.props.history.push({ pathname: "/customNewmessage", state: { data: this.state.templateType } })} secondary />}
                                            <Flex className="buttonContainer">
                                                <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                            </Flex>
                                        </Flex>

                                    </Flex>

                                </div>
                        }


                    </div>
                );

            }
            else if (this.state.page === "AudienceSelection") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half">
                                    <Flex column className="formContentContainer">
                                        <div>
                                            <h3>{this.localize("SelectTenantHeadingText")}</h3>
                                            <Dropdown
                                                placeholder={this.localize("SelectTenantHeadingText")}
                                                items={this.state.allTenantList}
                                                disabled={this.state.allTenantList.length > 0 ? false : true}
                                                value={this.state.selectedTenantName}
                                                onChange={(event, { value }) => this.selectTenant(value)}
                                                noResultsMessage={this.localize("NoMatchMessage")}
                                            />
                                        </div>
                                        {this.state.tenantId && <div>
                                            <h3>{this.localize("SendHeadingText")}</h3>
                                            <RadioGroup
                                                className="radioBtns"
                                                checkedValue={this.state.selectedRadioBtn}
                                                onCheckedValueChange={this.onGroupSelected}
                                                vertical={true}
                                                items={[
                                                    (((this.state.templateType === this.localize("ImageUpload")) || (this.state.templateType === this.localize("PDFUpload"))) && {
                                                        name: "teams",
                                                        key: "teams",
                                                        value: "teams",
                                                        label: this.localize("SendToGeneralChannel"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <Dropdown
                                                                        hidden={!this.state.teamsOptionSelected}
                                                                        placeholder={this.localize("SendToGeneralChannelPlaceHolder")}
                                                                        search
                                                                        multiple
                                                                        items={this.getItems()}
                                                                        value={this.state.selectedTeams}
                                                                        onChange={this.onTeamsChange}
                                                                        noResultsMessage={this.localize("NoMatchMessage")}
                                                                    />
                                                                </Flex>
                                                            )
                                                        },
                                                    }),
                                                    {
                                                        name: "rosters",
                                                        key: "rosters",
                                                        value: "rosters",
                                                        label: this.localize("SendToRosters"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <Dropdown
                                                                        hidden={!this.state.rostersOptionSelected}
                                                                        placeholder={this.localize("SendToRostersPlaceHolder")}
                                                                        search
                                                                        multiple
                                                                        items={this.getItems()}
                                                                        value={this.state.selectedRosters}
                                                                        onChange={this.onRostersChange}
                                                                        unstable_pinned={this.state.unstablePinned}
                                                                        noResultsMessage={this.localize("NoMatchMessage")}
                                                                    />
                                                                </Flex>
                                                            )
                                                        },
                                                    },
                                                    {
                                                        name: "allUsers",
                                                        key: "allUsers",
                                                        value: "allUsers",
                                                        label: this.localize("SendToAllUsers"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <div className={this.state.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                                                        <div className="noteText">
                                                                            <Text error content={this.localize("SendToAllUsersNote")} />
                                                                        </div>
                                                                    </div>
                                                                </Flex>
                                                            )
                                                        },
                                                    },
                                                    {
                                                        name: "groups",
                                                        key: "groups",
                                                        value: "groups",
                                                        label: this.localize("SendToGroups"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <div className={this.state.groupsOptionSelected && !this.state.groupAccess ? "" : "hide"}>
                                                                        <div className="noteText">
                                                                            <Text error content={this.localize("SendToGroupsPermissionNote")} />
                                                                        </div>
                                                                    </div>
                                                                    <Dropdown
                                                                        className="hideToggle"
                                                                        hidden={!this.state.groupsOptionSelected || !this.state.groupAccess}
                                                                        placeholder={this.localize("SendToGroupsPlaceHolder")}
                                                                        search={this.onGroupSearch}
                                                                        multiple
                                                                        loading={this.state.loading}
                                                                        loadingMessage={this.localize("LoadingText")}
                                                                        items={this.getGroupItems()}
                                                                        value={this.state.selectedGroups}
                                                                        onSearchQueryChange={this.onGroupSearchQueryChange}
                                                                        onChange={this.onGroupsChange}
                                                                        noResultsMessage={this.state.noResultMessage}
                                                                        unstable_pinned={this.state.unstablePinned}
                                                                    />
                                                                    <div className={this.state.groupsOptionSelected && this.state.groupAccess ? "" : "hide"}>
                                                                        <div className="noteText">
                                                                            <Text error content={this.localize("SendToGroupsNote")} />
                                                                        </div>
                                                                    </div>
                                                                </Flex>
                                                            )
                                                        },
                                                    },
                                                    {
                                                        name: "csv",
                                                        key: "csv",
                                                        value: "csv",
                                                        label: this.localize("SendToCSV"),
                                                        children: (Component, { name, ...props }) => {
                                                            return (
                                                                <Flex key={name} column>
                                                                    <Component {...props} />
                                                                    <div className={`csvFileUploadDiv ${this.state.selectedRadioBtn === "csv" ? "" : "hide"}`}>
                                                                        <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                                                            <Text size="medium" content={this.localize("UploadCSVFileText")} />
                                                                            <input type="file" id="upload" style={{ display: 'none' }} onChange={value => this.onFileChoose(value)} accept=".csv" />
                                                                            <Flex.Item push>
                                                                                <Button circular onClick={() => this.fileUpload()}
                                                                                    size="small"
                                                                                    icon={<FilesUploadIcon />}
                                                                                    title={this.localize("UploadText")}

                                                                                />
                                                                            </Flex.Item>
                                                                        </Flex>
                                                                        {this.state.csvLink && <div className="flexColumn">
                                                                            <Text size="small" content={this.localize("CsvUplodSuccessText")} color="green" />
                                                                            <a href={this.state.csvLink}>{this.localize("ViewFile")}</a>
                                                                        </div>
                                                                        }
                                                                    </div>
                                                                </Flex>
                                                            )
                                                        },
                                                    }
                                                ]}
                                            >
                                            </RadioGroup>
                                        </div>}


                                        <Flex hAlign="start">
                                            <h3><Checkbox
                                                className="ScheduleCheckbox"
                                                labelPosition="start"
                                                onClick={this.onScheduleSelected}
                                                label={this.localize("ScheduledSend")}
                                                checked={this.state.selectedSchedule}
                                                toggle
                                            /></h3>
                                        </Flex>
                                        <Text size="small" align="start" content={this.localize('ScheduledSendDescription')} />
                                        <Flex gap="gap.smaller" className="DateTimeSelector">
                                            <Datepicker
                                                disabled={!this.state.selectedSchedule}
                                                defaultSelectedDate={this.getDateObject(this.state.scheduledDate)}
                                                minDate={new Date()}
                                                inputOnly
                                                onDateChange={this.handleDateChange}
                                            />
                                            <Flex.Item shrink={true} size="1%">
                                                <Dropdown
                                                    placeholder="hour"
                                                    disabled={!this.state.selectedSchedule}
                                                    fluid={true}
                                                    items={hours}
                                                    defaultValue={this.getDateHour(this.state.scheduledDate)}
                                                    onChange={this.handleHourChange}
                                                />
                                            </Flex.Item>
                                            <Flex.Item shrink={true} size="1%">
                                                <Dropdown
                                                    placeholder="mins"
                                                    disabled={!this.state.selectedSchedule}
                                                    fluid={true}
                                                    items={minutes}
                                                    defaultValue={this.getDateMins(this.state.scheduledDate)}
                                                    onChange={this.handleMinsChange}
                                                />
                                            </Flex.Item>
                                        </Flex>
                                        <div className={this.state.futuredate && this.state.selectedSchedule ? "ErrorMessage" : "hide"}>
                                            <div className="noteText">
                                                <Text error content={this.localize('FutureDateError')} />
                                            </div>
                                        </div>
                                        <Flex hAlign="start">
                                            <h3><Checkbox
                                                className="Important"
                                                labelPosition="start"
                                                onClick={this.onImportantSelected}
                                                label={this.localize("Important")}
                                                checked={this.state.selectedImportant}
                                                toggle
                                            /></h3>
                                        </Flex>
                                        <Text size="small" align="start" content={this.localize('ImportantDescription')} />
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>
                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <Flex className="buttonContainer" gap="gap.medium">
                                    <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                    <Flex.Item push>
                                        <Button
                                            content="Schedule"
                                            disabled={this.isSaveBtnDisabled() || !this.state.selectedSchedule}
                                            onClick={this.onSchedule}
                                            primary={this.state.selectedSchedule} />
                                    </Flex.Item>
                                    <Button content={this.localize("SaveAsDraft")}
                                        disabled={this.isSaveBtnDisabled() || this.state.selectedSchedule}
                                        id="saveBtn"
                                        onClick={this.onSave}
                                        primary={!this.state.selectedSchedule} />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            } else {
                return (<div>Error</div>);
            }
        }
    }

    //get the next rounded up (ceil) date in minutes
    private getRoundedDate = (minutes: number, d = new Date()) => {

        let ms = 1000 * 60 * minutes; // convert minutes to ms
        let roundedDate = new Date(Math.ceil(d.getTime() / ms) * ms);

        return roundedDate
    }

    //get date object based on the string parameter
    private getDateObject = (datestring?: string) => {
        if (!datestring) {
            var TempDate = new Date(); //get current date
            TempDate.setTime(TempDate.getTime() + 86400000);
            return TempDate; //if date string is not provided, then return tomorrow rounded up next 5 minutes
        }
        return new Date(datestring); //if date string is provided, return current date object
    }

    //get the hour of the datestring
    private getDateHour = (datestring: string) => {
        if (!datestring) return "00";
        var thour = new Date(datestring).getHours().toString();
        return thour.padStart(2, "0");
    }

    //get the mins of the datestring
    private getDateMins = (datestring: string) => {
        if (!datestring) return "00";
        var tmins = new Date(datestring).getMinutes().toString();
        return tmins.padStart(2, "0");
    }

    //handles click on DatePicker to change the schedule date
    private handleDateChange = (e: any, v: any) => {
        var TempDate = v.value; //set the tempdate var with the value selected by the user
        TempDate.setMinutes(parseInt(this.state.DMYMins)); //set the minutes selected on minutes drop down 
        TempDate.setHours(parseInt(this.state.DMYHour)); //set the hour selected on hour drop down
        //set the state variables
        this.setState({
            scheduledDate: TempDate.toUTCString(), //updates the state string representation
            DMY: TempDate, //updates the date on the state
        });
    }

    //handles selection on the hour combo
    private handleHourChange = (e: any, v: any) => {
        var TempDate = this.state.DMY; //get the tempdate from the state
        TempDate.setHours(parseInt(v.value)); //set hour with the value select on the hour drop down
        //set state variables
        this.setState({
            scheduledDate: TempDate.toUTCString(), //updates the string representation 
            DMY: TempDate, //updates DMY
            DMYHour: v.value, //set the new hour value on the state
        });
    }

    //handles selection on the minutes combo
    private handleMinsChange = (e: any, v: any) => {
        var TempDate = this.state.DMY; //get the tempdate from the state
        TempDate.setMinutes(parseInt(v.value)); //set minutes with the value select on the minutes drop down
        //set state variables
        this.setState({
            scheduledDate: TempDate.toUTCString(), //updates the string representation 
            DMY: TempDate, //updates DMY
            DMYMins: v.value, //set the bew minutes on the state
        });
    }

    //handler for the Schedule Send checkbox
    private onScheduleSelected = () => {
        var TempDate = this.getRoundedDate(5, this.getDateObject()); //get the next day date rounded to the nearest hour/minute
        //set the state
        this.setState({
            selectedSchedule: !this.state.selectedSchedule,
            scheduledDate: TempDate.toUTCString(),
            DMY: TempDate
        });
    }

    // handler for the important message checkbox
    private onImportantSelected = () => {
        this.setState({
            selectedImportant: !this.state.selectedImportant
        });
    }

    private onGroupSelected = (event: any, data: any) => {
        this.setState({
            selectedRadioBtn: data.value,
            teamsOptionSelected: data.value === 'teams',
            rostersOptionSelected: data.value === 'rosters',
            groupsOptionSelected: data.value === 'groups',
            allUsersOptionSelected: data.value === 'allUsers' || data.value === "sistertenant",
            sisterTenantOptionSelected: data.value === 'sistertenant',
            csvOptionSelected: data.value === 'csv',
            selectedTeams: data.value === 'teams' ? this.state.selectedTeams : [],
            selectedTeamsNum: data.value === 'teams' ? this.state.selectedTeamsNum : 0,
            selectedRosters: data.value === 'rosters' ? this.state.selectedRosters : [],
            selectedRostersNum: data.value === 'rosters' ? this.state.selectedRostersNum : 0,
            selectedGroups: data.value === 'groups' ? this.state.selectedGroups : [],
            selectedGroupsNum: data.value === 'groups' ? this.state.selectedGroupsNum : 0,
        });
    }

    private isSaveBtnDisabled = () => {
        const teamsSelectionIsValid = (this.state.teamsOptionSelected && this.state.tenantId && (this.state.selectedTeamsNum !== 0)) || (!this.state.teamsOptionSelected);
        const rostersSelectionIsValid = (this.state.rostersOptionSelected && this.state.tenantId && (this.state.selectedRostersNum !== 0)) || (!this.state.rostersOptionSelected);
        const groupsSelectionIsValid = (this.state.groupsOptionSelected && this.state.tenantId && (this.state.selectedGroupsNum !== 0)) || (!this.state.groupsOptionSelected);
        const nothingSelected = (!this.state.teamsOptionSelected) && (!this.state.rostersOptionSelected) && (!this.state.groupsOptionSelected) && (!this.state.allUsersOptionSelected) && (!this.state.sisterTenantOptionSelected) && (!this.state.csvOptionSelected);
        const csvSelectionIsValid = (this.state.csvOptionSelected && !this.state.csvLink && this.state.tenantId)
        return (!teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected || csvSelectionIsValid)
    }

    private isNextBtnDisabled = () => {
        const title = this.state.title;
        const imageLink = this.state.imageLink
        if (this.state.templateType === this.localize("Q&AUpload")) {
            const lastIndex = this.state.questionAnswer.length - 1;
            const question = (this.state.questionAnswer.length > 0) && this.state.questionAnswer[lastIndex].question;
            if (this.state.questionAnswer.length > 0 && this.state.questionAnswer[lastIndex].questionType === this.localize("MCQ")) {
                const answer1 = this.state.questionAnswer[lastIndex].answer[0];
                const answer2 = this.state.questionAnswer[lastIndex].answer[1];
                const answer3 = this.state.questionAnswer[lastIndex].answer[2];
                const answer4 = this.state.questionAnswer[lastIndex].answer[3];
                return !(title && (this.state.addQuestionError === "") && (this.state.questionAnswer.length > 0) && question && answer1 && answer2 && answer3 && answer4);
            }
            else {
                return !(title && (this.state.addQuestionError === "") && (this.state.questionAnswer.length > 0) && question);
            }

        }
        else if (this.state.templateType === this.localize("EmailUpload")) {
            return !(title && (this.state.errorEmailUrlMessage === "") && imageLink && this.state.emailFileTitle);
        }

        else {
            const btnArraylastIndex = this.state.values.length - 1;
            const btnTitle = (this.state.values.length > 0) && this.state.values[btnArraylastIndex].title;
            const btnLink = (this.state.values.length > 0) && this.state.values[btnArraylastIndex].url;
            return !(title && (this.state.errorButtonUrlMessage === "") && imageLink && ((this.state.values.length > 0) ? (btnTitle && btnLink) : true));
        }

    }

    private getItems = () => {
        const resultedTeams: dropdownItem[] = [];
        if (this.state.teams) {
            let remainingUserTeams = this.state.teams;
            if (this.state.selectedRadioBtn !== "allUsers") {
                if (this.state.selectedRadioBtn === "teams") {
                    this.state.teams.filter(x => this.state.selectedTeams.findIndex(y => y.team.id === x.id) < 0);
                }
                else if (this.state.selectedRadioBtn === "rosters") {
                    this.state.teams.filter(x => this.state.selectedRosters.findIndex(y => y.team.id === x.id) < 0);
                }
            }
            remainingUserTeams.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id
                    }
                });
            });
        }
        return resultedTeams;
    }

    private static MAX_SELECTED_TEAMS_NUM: number = 20;

    private onTeamsChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedTeams: itemsData.value,
            selectedTeamsNum: itemsData.value.length,
            selectedRosters: [],
            selectedRostersNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0
        })
    }

    private onRostersChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedRosters: itemsData.value,
            selectedRostersNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0
        })
    }

    private onGroupsChange = (event: any, itemsData: any) => {
        this.setState({
            selectedGroups: itemsData.value,
            selectedGroupsNum: itemsData.value.length,
            groups: [],
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedRosters: [],
            selectedRostersNum: 0
        })
    }

    private onGroupSearch = (itemList: any, searchQuery: string) => {
        const result = itemList.filter(
            (item: { header: string; content: string; }) => (item.header && item.header.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1) ||
                (item.content && item.content.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1),
        )
        return result;
    }

    private onGroupSearchQueryChange = async (event: any, itemsData: any) => {

        if (!itemsData.searchQuery) {
            this.setState({
                groups: [],
                noResultMessage: "",
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
            this.setState({
                loading: false,
                noResultMessage: this.localize("NoMatchMessage"),
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
            // handle event trigger on item select.
            const result = itemsData.items && itemsData.items.find(
                (item: { header: string; }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
            )
            if (result) {
                return;
            }

            this.setState({
                loading: true,
                noResultMessage: "",
            });

            try {
                const query = encodeURIComponent(itemsData.searchQuery);
                const response = await searchGroups(query, this.state.tenantId);
                console.log("m365 group search", response)
                this.setState({
                    groups: response.data,
                    loading: false,
                    noResultMessage: this.localize("NoMatchMessage")
                });
            }
            catch (error) {
                return error;
            }
        }
    }

    //called when the user clicks to schedule the message
    private onSchedule = () => {
        var Today = new Date(); //today date
        var Scheduled = new Date(this.state.DMY); //scheduled date

        //only allow the save when the scheduled date is 30 mins in the future, if that is the case calls the onSave function
        if (Scheduled.getTime() > Today.getTime() + 1800000) { this.onSave() }
        else {
            //set the state to indicate future date error
            //if futuredate is true, an error message is shown right below the date selector
            this.setState({
                futuredate: true
            })
        }
    }

    //called to save the draft
    private onSave = () => {
        const selectedTeams: string[] = [];
        const selctedRosters: string[] = [];
        const selectedGroups: string[] = [];
        this.state.selectedTeams.forEach(x => selectedTeams.push(x.team.id));
        this.state.selectedRosters.forEach(x => selctedRosters.push(x.team.id));
        this.state.selectedGroups.forEach(x => selectedGroups.push(x.team.id));

        const draftMessage: IDraftMessage = {
            id: this.state.messageId,
            title: this.state.title,
            imageLink: this.state.imageLink,
            summary: (this.state.templateType === this.localize("Q&AUpload")) ? (this.state.questionAnswer.length > 0) ? JSON.stringify(this.state.questionAnswer) : "" : this.state.summary,
            author: this.state.author,
            buttonTitle: this.state.btnTitle,
            buttonLink: this.state.btnLink,
            teams: selectedTeams,
            rosters: selctedRosters,
            groups: selectedGroups,
            allUsers: this.state.allUsersOptionSelected,
            isScheduled: this.state.selectedSchedule,
            isImportant: this.state.selectedImportant,
            ScheduledDate: new Date(this.state.scheduledDate),
            Buttons: JSON.stringify(this.state.values),
            TemplateType: this.state.templateType,
            SendTypeId: (this.state.selectedRadioBtn === "teams") ? "1" : (this.state.selectedRadioBtn === "rosters") ? "2" : (this.state.selectedRadioBtn === "allUsers") ? "3" : (this.state.selectedRadioBtn === "groups") ? "4" : (this.state.selectedRadioBtn === "sistertenant") ? "5" : "6",
            CsvLink: this.state.csvLink,
            EmailBody: this.state.emailBodyContent,
            EmailTitle: this.state.emailFileTitle,
            AdaptiveCardContent: JSON.stringify(this.card),
            TenantId: this.state.tenantId,
            TenantName: this.state.selectedTenantName,
            AuthorTeamId: this.state.AuthorTeamId,
            AuthorTeamName: this.state.AuthorTeamName,
            AuthorChannelId: this.state.AuthorChannelId,
            AuthorChannelName: this.state.AuthorChannelName,
        };


        if (this.state.exists) {
            this.editDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        } else {
            this.postDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        }
    }

    private editDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            console.log("edit draft messsage", draftMessage);
            await updateDraftNotification(draftMessage);
        } catch (error) {
            return error;
        }
    }

    private postDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            console.log("draft messsage", draftMessage);
            await createDraftNotification(draftMessage);
        } catch (error) {
            throw error;
        }
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    private onNext = (event: any) => {
        this.setState({
            page: "AudienceSelection"
        }, () => {
            this.updateCard();
        });
    }

    private onBack = (event: any) => {
        this.setState({
            page: "CardCreation"
        }, () => {
            this.updateCard();
        });
    }

    private onTitleChanged = (event: any) => {
        let showDefaultCard = (!event.target.value && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        if (this.state.templateType === this.localize("ImageUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsName(this.card, teamsName)
            setCardTitle(this.card, event.target.value);
            setCardImageLink(this.card, this.state.imageLink);
            setCardSummary(this.card, this.state.summary);
            setCardAuthor(this.card, this.state.author);
            setCardBtns(this.card, this.state.values);
        }
        else if (this.state.templateType === this.localize("PDFUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNamePDFUpload(this.card, teamsName)
            setCardTitlePDFUpload(this.card, event.target.value);
            if (this.state.imageLink !== "") {
                setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                let pdfLink = "[View PDF](" + this.state.imageLink + ")"
                setCardPdfNamePDFUpload(this.card, pdfLink)
            }
            setCardSummaryPDFUpload(this.card, this.state.summary);
            setCardAuthorPDFUpload(this.card, this.state.author);
            setCardBtnsPDFUpload(this.card, this.state.values);
        }
        else if (this.state.templateType === this.localize("Q&AUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNameQuestionAnswere(this.card, teamsName)
            setCardTitleQuestionAnswer(this.card, event.target.value);
            setCardAuthorQuestionAnswer(this.card, this.state.author);
        }
        else {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            let EmailLink = (this.state.imageLink !== "") ? "[" + this.state.emailFileTitle + "](" + this.state.imageLink + ")" : ""
            setCardEmailTemplate(this.card, teamsName, this.state.emailBodyContent, event.target.value, this.state.author, EmailLink, this.state.summary);

        }


        this.setState({
            title: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onImageLinkChanged = (event: any) => {
        let url = event.target.value.toLowerCase();
        if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,")) || (url.startsWith("data:application/pdf"))))) {
            this.setState({
                errorImageUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorImageUrlMessage: ""
            });
        }

        let showDefaultCard = (!this.state.title && !event.target.value && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        if (this.state.templateType === this.localize("ImageUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsName(this.card, teamsName)
            setCardTitle(this.card, this.state.title);
            setCardImageLink(this.card, event.target.value);
            setCardSummary(this.card, this.state.summary);
            setCardAuthor(this.card, this.state.author);
            setCardBtns(this.card, this.state.values);
        }
        this.setState({
            imageLink: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onSummaryChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !event.target.value && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        if (this.state.templateType === this.localize("ImageUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsName(this.card, teamsName)
            setCardTitle(this.card, this.state.title);
            setCardImageLink(this.card, this.state.imageLink);
            setCardSummary(this.card, event.target.value);
            setCardAuthor(this.card, this.state.author);
            setCardBtns(this.card, this.state.values);
        }
        else if (this.state.templateType === this.localize("PDFUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNamePDFUpload(this.card, teamsName)
            setCardTitlePDFUpload(this.card, this.state.title);
            if (this.state.imageLink !== "") {
                setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                let pdfLink = "[View PDF](" + this.state.imageLink + ")"
                setCardPdfNamePDFUpload(this.card, pdfLink)
            }
            setCardSummaryPDFUpload(this.card, event.target.value);
            setCardAuthorPDFUpload(this.card, this.state.author);
            setCardBtnsPDFUpload(this.card, this.state.values);
        }
        else if (this.state.templateType === this.localize("EmailUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            let EmailLink = (this.state.imageLink !== "") ? "[" + this.state.emailFileTitle + "](" + this.state.imageLink + ")" : ""
            setCardEmailTemplate(this.card, teamsName, this.state.emailBodyContent, this.state.title, this.state.author, EmailLink, event.target.value)
        }
        this.setState({
            summary: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    //if the author changes, updates the card with appropriate values
    private onAuthorChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !event.target.value && !this.state.btnTitle && !this.state.btnLink);
        if (this.state.templateType === this.localize("ImageUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsName(this.card, teamsName)
            setCardTitle(this.card, this.state.title);
            setCardImageLink(this.card, this.state.imageLink);
            setCardSummary(this.card, this.state.summary);
            setCardAuthor(this.card, event.target.value);
            setCardBtns(this.card, this.state.values);
        }
        else if (this.state.templateType === this.localize("PDFUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNamePDFUpload(this.card, teamsName)
            setCardTitlePDFUpload(this.card, this.state.title);
            if (this.state.imageLink !== "") {
                setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                let pdfLink = "[View PDF](" + this.state.imageLink + ")"
                setCardPdfNamePDFUpload(this.card, pdfLink)
            }
            setCardSummaryPDFUpload(this.card, this.state.summary);
            setCardAuthorPDFUpload(this.card, event.target.value);
            setCardBtnsPDFUpload(this.card, this.state.values);
        }
        else if (this.state.templateType === this.localize("Q&AUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNameQuestionAnswere(this.card, teamsName)
            setCardTitleQuestionAnswer(this.card, this.state.title);
            setCardAuthorQuestionAnswer(this.card, event.target.value);
            // setCardBtnsQuestionAnswer(this.card, this.state.values);
        }
        else {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            let EmailLink = (this.state.imageLink !== "") ? "[" + this.state.emailFileTitle + "](" + this.state.imageLink + ")" : ""
            setCardEmailTemplate(this.card, teamsName, this.state.emailBodyContent, this.state.title, event.target.value, EmailLink, this.state.summary)
        }
        this.setState({
            author: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private EmailTitle = (event: any) => {
        this.setState({
            emailFileTitle: event.target.value
        }, () => {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            let EmailLink = (this.state.imageLink !== "") ? "[" + this.state.emailFileTitle + "](" + this.state.imageLink + ")" : ""
            setCardEmailTemplate(this.card, teamsName, this.state.emailBodyContent, this.state.title, this.state.author, EmailLink, this.state.summary)
            this.updateCard();
        });
    }

    // private function to create the buttons UI
    private createUI() {
        if (this.state.values.length > 0) {
            return this.state.values.map((el, i) =>
                <Flex gap="gap.smaller" vAlign="center" styles={{ marginTop: "5px" }}>
                    <Input className="inputField"
                        fluid
                        value={el.title || ''}
                        placeholder={this.localize("ButtonTitle")}
                        onChange={this.handleChangeName.bind(this, i)}
                        autoComplete="off"
                    />
                    <Input className="inputField"
                        fluid
                        value={el.url || ''}
                        placeholder={this.localize("ButtonURL")}
                        onChange={this.handleChangeLink.bind(this, i)}
                        error={!(this.state.errorButtonUrlMessage === "")}
                        autoComplete="off"
                    />
                    <Button
                        circular
                        size="small"
                        icon={<TrashCanIcon />}
                        onClick={this.removeClick.bind(this, i)}
                        title={this.localize("Delete")}
                    />
                </Flex>
            )
        } else {
            return (
                < Flex >
                    <Text size="small" content={this.localize("NoButtons")} />
                </Flex>
            )
        }
    }

    //private function to add a new button to the adaptive card
    private addClick() {
        const item =
        {
            type: "Action.OpenUrl",
            title: "",
            url: ""
        };

        this.setState({
            values: [...this.state.values, item]
        });
    }

    //private function to remove a button from the adaptive card
    private removeClick(i: any) {
        let values = [...this.state.values];
        values.splice(i, 1);
        this.setState({ values });

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && values.length == 0);

        if (this.state.templateType === this.localize("ImageUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsName(this.card, teamsName)
            setCardTitle(this.card, this.state.title);
            setCardImageLink(this.card, this.state.imageLink);
            setCardSummary(this.card, this.state.summary);
            setCardAuthor(this.card, this.state.author);
        }
        else if (this.state.templateType === this.localize("PDFUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNamePDFUpload(this.card, teamsName)
            setCardTitlePDFUpload(this.card, this.state.title);
            if (this.state.imageLink !== "") {
                setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                let pdfLink = "[View PDF](" + this.state.imageLink + ")"
                setCardPdfNamePDFUpload(this.card, pdfLink)
            }
            setCardSummaryPDFUpload(this.card, this.state.summary);
            setCardAuthorPDFUpload(this.card, this.state.author);
        }
        else if (this.state.templateType === this.localize("Q&AUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNameQuestionAnswere(this.card, teamsName)
            setCardTitleQuestionAnswer(this.card, this.state.title);
            setCardAuthorQuestionAnswer(this.card, this.state.author);
        }
        if (values.length > 0) { //only if there are buttons created
            //update the adaptive card
            if (this.state.templateType === this.localize("ImageUpload")) {
                setCardBtns(this.card, values);
            }
            else if (this.state.templateType === this.localize("PDFUpload")) {
                setCardBtnsPDFUpload(this.card, values);
            }
            // else if (this.state.templateType === this.localize("Q&AUpload")) {
            //     setCardBtnsQuestionAnswer(this.card, values);
            // }
            this.setState({
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
            delete this.card.actions;
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        };
    }

    //private function to deal with changes in the button names
    private handleChangeName(i: any, event: any) {
        let values = [...this.state.values];
        values[i].title = event.target.value;
        this.setState({ values });

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !event.target.value && values.length == 0);
        if (this.state.templateType === this.localize("ImageUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsName(this.card, teamsName)
            setCardTitle(this.card, this.state.title);
            setCardImageLink(this.card, this.state.imageLink);
            setCardSummary(this.card, this.state.summary);
            setCardAuthor(this.card, this.state.author);
        }
        else if (this.state.templateType === this.localize("PDFUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNamePDFUpload(this.card, teamsName)
            setCardTitlePDFUpload(this.card, this.state.title);
            if (this.state.imageLink !== "") {
                setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                let pdfLink = "[View PDF](" + this.state.imageLink + ")"
                setCardPdfNamePDFUpload(this.card, pdfLink)
            }
            setCardSummaryPDFUpload(this.card, this.state.summary);
            setCardAuthorPDFUpload(this.card, this.state.author);
        }
        else if (this.state.templateType === this.localize("Q&AUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNameQuestionAnswere(this.card, teamsName)
            setCardTitleQuestionAnswer(this.card, this.state.title);
            setCardAuthorQuestionAnswer(this.card, this.state.author);
        }
        if (values.length > 0) { //only if there are buttons created
            //update the adaptive card
            if (this.state.templateType === this.localize("ImageUpload")) {
                setCardBtns(this.card, values);
            }
            else if (this.state.templateType === this.localize("PDFUpload")) {
                setCardBtnsPDFUpload(this.card, values);
            }
            // else if (this.state.templateType === this.localize("Q&AUpload")) {
            //     setCardBtnsQuestionAnswer(this.card, values);
            // }
            this.setState({
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        };
    }

    //private function to deal with changes in the button links/urls
    private handleChangeLink(i: any, event: any) {
        let values = [...this.state.values];
        values[i].url = event.target.value;
        this.setState({ values });

        //set the error message if the links have wrong values
        //alert(values.findIndex(element => element.includes("https://")));
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
        }

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !event.target.value && values.length == 0);
        if (this.state.templateType === this.localize("ImageUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsName(this.card, teamsName)
            setCardTitle(this.card, this.state.title);
            setCardImageLink(this.card, this.state.imageLink);
            setCardSummary(this.card, this.state.summary);
            setCardAuthor(this.card, this.state.author);
        }
        else if (this.state.templateType === this.localize("PDFUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNamePDFUpload(this.card, teamsName)
            setCardTitlePDFUpload(this.card, this.state.title);
            if (this.state.imageLink !== "") {
                setCardImageLinkPDFUpload(this.card, pdfImgUrl);
                let pdfLink = "[View PDF](" + this.state.imageLink + ")"
                setCardPdfNamePDFUpload(this.card, pdfLink)
            }
            setCardSummaryPDFUpload(this.card, this.state.summary);
            setCardAuthorPDFUpload(this.card, this.state.author);
        }
        else if (this.state.templateType === this.localize("Q&AUpload")) {
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardTeamsNameQuestionAnswere(this.card, teamsName)
            setCardTitleQuestionAnswer(this.card, this.state.title);
            setCardAuthorQuestionAnswer(this.card, this.state.author);
        }
        if (values.length > 0) {
            //update the adaptive card
            if (this.state.templateType === this.localize("ImageUpload")) {
                setCardBtns(this.card, values);
            }
            else if (this.state.templateType === this.localize("PDFUpload")) {
                setCardBtnsPDFUpload(this.card, values);
            }

            this.setState({
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        };
    }


    // private function to create the Question Answer UI
    private createQuestionAnswerUI() {
        if (this.state.questionAnswer.length > 0) {
            return this.state.questionAnswer.map((el, i) =>
                <Flex gap="gap.smaller" vAlign="center" styles={{ marginTop: "5px" }}>
                    <div style={{ width: "100%" }}>
                        <Input className="inputField"
                            fluid
                            value={el.question || ''}
                            placeholder={this.localize("Question")}
                            onChange={this.handleChangeQuestion.bind(this, i)}
                            autoComplete="off"
                        />
                        {(el.questionType === this.localize("MCQ")) && <div>
                            {/* {this.MCQAnswerUI(i)} */}
                            <Flex gap="gap.smaller" vAlign="center" styles={{ marginTop: "5px" }}>
                                <Input className="inputField"
                                    fluid
                                    value={el.answer[0] || ''}
                                    placeholder={this.localize("Option") + "1"}
                                    onChange={(event: any) => this.handleChangeMCQAnswer(event.target.value, i, 0)}
                                    autoComplete="off"
                                />
                                <Input className="inputField"
                                    fluid
                                    value={el.answer[1] || ''}
                                    placeholder={this.localize("Option") + "2"}
                                    onChange={(event: any) => this.handleChangeMCQAnswer(event.target.value, i, 1)}
                                    autoComplete="off"
                                />
                            </Flex>
                            <Flex gap="gap.smaller" vAlign="center" styles={{ marginTop: "5px" }}>
                                <Input className="inputField"
                                    fluid
                                    value={el.answer[2] || ''}
                                    placeholder={this.localize("Option") + "3"}
                                    onChange={(event: any) => this.handleChangeMCQAnswer(event.target.value, i, 2)}
                                    autoComplete="off"
                                />
                                <Input className="inputField"
                                    fluid
                                    value={el.answer[3] || ''}
                                    placeholder={this.localize("Option") + "4"}
                                    onChange={(event: any) => this.handleChangeMCQAnswer(event.target.value, i, 3)}
                                    autoComplete="off"
                                />
                            </Flex>
                        </div>}
                    </div>

                    <Button
                        circular
                        size="small"
                        icon={<TrashCanIcon />}
                        onClick={this.removeQuestions.bind(this, i)}
                        title={this.localize("Delete")}
                    />
                </Flex>
            )
        } else {
            return (
                < Flex >
                    <Text size="small" content={this.localize("NoQuestions")} />
                </Flex>
            )
        }
    }

    // private MCQAnswerUI(i: any) {
    //     var MCQArray = ["", "", "", ""];
    //     var a = MCQArray.map((ele: any, j: any) => {
    //         return <div>
    //             <Input type='text'
    //                 placeholder={this.localize("Option") + ` ${j + 1}`}
    //                 onChange={(event) => this.handleChangeMCQAnswer(event.target.value, i, j)} />
    //         </div>
    //     })
    //     return a;
    // }


    //private function to add a new question answer to the adaptive card

    private addQuestionAnswer() {
        const lastIndex = this.state.questionAnswer.length - 1;
        const item =
        {
            questionType: this.state.questionTypeSelectedValue,
            question: "",
            answer: ["", "", "", ""],
            type: (this.state.questionTypeSelectedValue === this.localize("DescriptiveQuestion")) ? "TextBlock" : "Input.ChoiceSet"
        };
        if (this.state.questionAnswer.length > 0) {
            if (this.state.questionAnswer[lastIndex].question !== "") {
                if (this.state.questionAnswer[lastIndex].questionType === this.localize("MCQ")) {
                    if ((this.state.questionAnswer[lastIndex].answer[0] !== "" && this.state.questionAnswer[lastIndex].answer[1] !== "" && this.state.questionAnswer[lastIndex].answer[2] !== "" && this.state.questionAnswer[lastIndex].answer[3] !== "")) {

                        this.setState({
                            questionAnswer: [...this.state.questionAnswer, item],
                        });
                    }
                    else {
                        this.setState({
                            addQuestionError: this.localize("MCQAnswerError")
                        });
                    }
                }
                else {
                    this.setState({
                        questionAnswer: [...this.state.questionAnswer, item]
                    });
                }
            }
            else {
                this.setState({
                    addQuestionError: this.localize("QuestionError")
                });
            }

        } else {
            this.setState({
                questionAnswer: [...this.state.questionAnswer, item]
            });
        }
    }

    private mcqLimitationCheck() {
        const mcqCount = this.state.questionAnswer.filter((e) => e.questionType === this.localize("MCQ"));
        if (mcqCount.length < 2) {
            this.setState({
                questionTypeSelectedValueDisable: false
            }, () => {
                this.addQuestionAnswer()
            })
        }
        else {
            this.setState({
                questionTypeSelectedValueDisable: true,
                questionTypeSelectedValue: this.localize("DescriptiveQuestion"),
            }, () => {
                this.addQuestionAnswer()
            })
        }
    }


    //private function to deal with changes in the questions
    private handleChangeQuestion(i: any, event: any) {
        let questionAnswer = [...this.state.questionAnswer];
        questionAnswer[i].question = event.target.value;
        this.setState({ questionAnswer });
        this.setState({ addQuestionError: "" });

        const showDefaultCard = (!this.state.title && !this.state.author && !event.target.value && questionAnswer.length == 0);

        if (questionAnswer.length > 0) { //only if there are questions created
            let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
            setCardPartQuestionAnswer(this.card, this.state.questionAnswer, this.localize, this.state.title, this.state.author, teamsName) //update the adaptive card

            this.setState({
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        };
    }


    handleChangeMCQAnswer = (event: any, i: any, j: any) => {
        let questionAnswer = [...this.state.questionAnswer];
        this.setState({ addQuestionError: "" });
        questionAnswer[i].answer[j] = event
        this.setState({ questionAnswer }, () => {
            // console.log("question answer", this.state.questionAnswer)
            const showDefaultCard = (!this.state.title && !this.state.author && !event && questionAnswer.length == 0);

            if (questionAnswer.length > 0) { //only if there are questions created
                let teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
                setCardPartQuestionAnswer(this.card, this.state.questionAnswer, this.localize, this.state.title, this.state.author, teamsName); //update the adaptive card
                this.setState({
                    card: this.card
                }, () => {
                    if (showDefaultCard) {
                        this.setDefaultCard(this.card);
                    }
                    this.updateCard();
                });
            } else {
                delete this.card.actions;
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            };
        });
    }



    //private function to remove a question from the adaptive card
    private removeQuestions(i: any) {

        let questionAnswer = [...this.state.questionAnswer];
        questionAnswer.splice(i, 1);
        this.setState({ questionAnswer }, () => {
            const mcqCount = this.state.questionAnswer.filter((e) => e.questionType === this.localize("MCQ"));
            if (mcqCount.length < 2) {
                this.setState({
                    questionTypeSelectedValueDisable: false
                })
            }
        });
        this.setState({
            addQuestionError: ""
        });


        const showDefaultCard = (!this.state.title && !this.state.author && questionAnswer.length == 0);
        const teamsName = this.state.AuthorTeamName + " >> " + this.state.AuthorChannelName
        setCardPartQuestionAnswer(this.card, questionAnswer, this.localize, this.state.title, this.state.author, teamsName); //update the adaptive card
        this.setState({
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }




    private updateCard = () => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(this.state.card);
        const renderedCard = adaptiveCard.render();
        const container = document.getElementsByClassName('adaptiveCardContainer')[0].firstChild;
        if (container != null) {
            container.replaceWith(renderedCard);
        } else {
            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
        }
        const link = this.state.btnLink;
        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
    }
}

const newMessageWithTranslation = withTranslation()(NewMessage);
export default newMessageWithTranslation;
