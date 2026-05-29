// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from '../configVariables';
import {  Button } from "@fluentui/react-northstar";

export interface IConfigState {
    url: string;
}

class Configuration extends React.Component<{}, IConfigState> {
    constructor(props: {}) {
        super(props);
        this.state = {
            url: getBaseUrl() + "/messages?locale={locale}"
        }
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
            microsoftTeams.settings.setSettings({
                entityId: "Company_Communicator_App",
                contentUrl: this.state.url,
                suggestedDisplayName: "Company Communicator",
            });
            saveEvent.notifySuccess();
        });

        microsoftTeams.settings.setValidityState(true);

    }


    public render(): JSX.Element {
        return (
            <div className="configContainer">
                <h3>Please click Save to get started.</h3>
                {/* <Button primary onClick={()=>this.submit()}>Configure</Button> */}
            </div>
        );
    }

}

export default Configuration;
