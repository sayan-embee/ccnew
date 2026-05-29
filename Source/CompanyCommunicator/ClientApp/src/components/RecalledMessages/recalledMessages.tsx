// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { connect } from 'react-redux';
import { withTranslation, WithTranslation } from "react-i18next";
import { TooltipHost } from 'office-ui-fabric-react';
import { Loader, List, Flex, Text, AcceptIcon, CloseIcon, ExclamationCircleIcon, EmailIcon } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";

import { selectMessage, getRecalledMessagesList, getDraftMessagesList,getMessagesList } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import Overflow from '../OverFlow/recalledMessageOverflow';
import './recalledMessages.scss';
import { TFunction } from "i18next";
import { formatNumber } from '../../i18n';

export interface ITaskInfo {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    card?: string;
    fallbackUrl?: string;
    completionBotId?: string;
}

export interface IRecalledMessages {
    title: string;
    sentDate: string;
    recipients: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
}

export interface IRecalledMessagesProps extends WithTranslation {
    messagesList: IRecalledMessages[];
    selectMessage?: any;
    getRecalledMessagesList?: any;
    getMessagesList?:any;
}

export interface IRecalledMessageState {
    message: IRecalledMessages[];
    loader: boolean;
}

class RecalledMessages extends React.Component<IRecalledMessagesProps, IRecalledMessageState> {
    readonly localize: TFunction;
    private interval: any;
    private isOpenTaskModuleAllowed: boolean;
    constructor(props: IRecalledMessagesProps) {
        super(props);
        this.localize = this.props.t;
        this.isOpenTaskModuleAllowed = true;
        console.log("Recalled Message : ",this.props.messagesList)
        this.state = {
            message: this.props.messagesList,
            loader: true,
        };
        this.escFunction = this.escFunction.bind(this);
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        this.props.getRecalledMessagesList();
        document.addEventListener("keydown", this.escFunction, false);
        this.interval = setInterval(() => {
            this.props.getRecalledMessagesList();
        }, 60000);
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
        clearInterval(this.interval);
    }

    public componentWillReceiveProps(nextProps: any) {
        if (this.props !== nextProps) {
            this.setState({
                message: nextProps.messagesList,
                loader: false
            });
        }
    }

    public render(): JSX.Element {
        let keyCount = 0;
        const processItem = (message: any) => {
            keyCount++;
            const out = {
                key: keyCount,
                content: this.messageContent(message),
                onClick: (): void => {
                    let url = getBaseUrl() + "/viewstatus/" + message.id + "?locale={locale}&openfrom=recall" ;
                    console.log("sayan url", url)
                    this.onOpenTaskModule(null, url, this.localize("ViewRecallStatus"));
                },
                styles: { margin: '0.2rem 0.2rem 0 0' },
            };
            return out;
        };

        const label = this.processLabels();
        const outList = this.state.message.map(processItem);
        const allMessages = [...label, ...outList];

        if (this.state.loader) {
            return (
                <Loader />
            );
        } else if (this.state.message.length === 0) {
            return (<div className="results">{this.localize("EmptyRecallMessages")}</div>);
        }
        else {
            return (
                <List selectable items={allMessages} className="list" />
            );
        }
    }

    private processLabels = () => {
        const out = [{
            key: "labels",
            content: (
                <Flex vAlign="center" fill gap="gap.small">
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} grow={1} >
                        <Text
                            truncated
                            weight="bold"
                            content={this.localize("TitleText")}
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '15%' }}>
                        <Text truncated
                            weight="bold"
                            content={this.localize("Tenant")}>
                            </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '30%' }}>
                        <Text>
                            </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} shrink={false}>
                        <Text
                            truncated
                            content={this.localize("RecalledStatus")}
                            weight="bold"
                        >
                        </Text>
                    </Flex.Item>                    
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '15%' }} >
                        <Text
                            truncated
                            content={this.localize("Recalled")}
                            weight="bold"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item shrink={0} >
                        <Overflow title="" />
                    </Flex.Item>
                </Flex>
            ),
            styles: { margin: '0.2rem 0.2rem 0 0' },
        }];
        return out;
    }

    private renderSendingText = (message: any) => {
        var text = "";
        switch (message.status) {
            case "Queued":
                text = this.localize("Queued");
                break;
            case "Recalling":
                let sentCount =
                    (message.succeeded ? message.succeeded : 0) +
                    (message.failed ? message.failed : 0) +
                    (message.unknown ? message.unknown : 0);

                text = this.localize("RecallingMessages", {"SentCount": formatNumber(sentCount), "TotalCount": formatNumber(message.totalMessageCount) });
                break;
            case "Recalled":
            case "Failed":
                text = "";
        }

        return (<Text truncated content={text} />);
    }

    private messageContent = (message: any) => {
        console.log("sayan",message)
        return (
            <Flex className="listContainer" vAlign="center" fill gap="gap.small">
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} grow={1}>
                    <Text
                        truncated
                        content={message.title}
                    >
                    </Text>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '15%' }} grow={1}>
                    <Text
                        truncated
                        content={message.tenantName}
                    >
                    </Text>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '30%' }}>
                    {this.renderSendingText(message)}
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} shrink={false}>
                    <div>
                        <TooltipHost content={this.props.t("TooltipTotalMessage")} calloutProps={{ gapSpace: 0 }}>
                            <EmailIcon xSpacing="after" className="succeeded" outline />
                            <span className="semiBold marginRight">{formatNumber(message.totalMessageCount)}</span>
                        </TooltipHost>
                        <TooltipHost content={this.props.t("TooltipSuccess")} calloutProps={{ gapSpace: 0 }} className="marginLeft">
                            <AcceptIcon xSpacing="after" className="succeeded" outline />
                            <span className="semiBold">{formatNumber(message.succeeded)}</span>
                        </TooltipHost>
                        <TooltipHost content={this.props.t("TooltipFailure")} calloutProps={{ gapSpace: 0 }}>
                            <CloseIcon xSpacing="both" className="failed" outline />
                            <span className="semiBold">{formatNumber(message.failed)}</span>
                        </TooltipHost>
                        {
                            message.unknown &&
                            <TooltipHost content="Unknown" calloutProps={{ gapSpace: 0 }}>
                                <ExclamationCircleIcon xSpacing="both" className="unknown" outline />
                                <span className="semiBold">{formatNumber(message.unknown)}</span>
                            </TooltipHost>
                        }
                    </div>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '15%' }} >
                    <Text
                        truncated
                        className="semiBold"
                        content={message.sentDate}
                    />
                </Flex.Item>
                <Flex.Item shrink={0}>
                    <Overflow message={message} title="" />
                </Flex.Item>
            </Flex>
        );
    }

    private escFunction = (event: any) => {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    public onOpenTaskModule = (event: any, url: string, title: string) => {
        if (this.isOpenTaskModuleAllowed) {
            this.isOpenTaskModuleAllowed = false;
            let taskInfo: ITaskInfo = {
                url: url,
                title: title,
                height: 530,
                width: 1000,
                fallbackUrl: url,
            }

            let submitHandler = (err: any, result: any) => {
                this.isOpenTaskModuleAllowed = true;
            };

            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        }
    }
}

const mapStateToProps = (state: any) => {
    return { messagesList: state.recalledMessagesList };
}

const RecalledMessagesWithTranslation = withTranslation()(RecalledMessages);
export default connect(mapStateToProps, { selectMessage, getRecalledMessagesList, getDraftMessagesList,getMessagesList })(RecalledMessagesWithTranslation);