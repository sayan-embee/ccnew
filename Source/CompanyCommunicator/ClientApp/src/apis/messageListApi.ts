// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import axios from './axiosJWTDecorator';
import { getBaseUrl } from '../configVariables';

let baseAxiosUrl = getBaseUrl() + '/api';

export const getSentNotifications = async (): Promise<any> => {
    //let url = baseAxiosUrl + "/sentnotifications";
    let url = baseAxiosUrl + "/sentnotifications/Updated2";
    return await axios.get(url);
}

export const getScheduledNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/scheduled";
    return await axios.get(url);
}

export const getDraftNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.get(url);
}

export const verifyGroupAccess = async (): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/verifyaccess";
    return await axios.get(url, false);
}

export const getGroups = async (id: number, tenantId: any): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/" + id + "/"+tenantId;
    return await axios.get(url);
}

export const searchGroups = async (query: string, tenantId: any): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/search/" + query + "/" + tenantId;
    return await axios.get(url);
}

export const exportNotification = async(payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/exportnotification/export";
    return await axios.put(url, payload);
}

export const getSentNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications/" + id;
    return await axios.get(url);
}

export const getRecallNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/recallSentNotifications/" + id;
    return await axios.get(url);
}

export const getDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.get(url);
}


export const deleteDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.delete(url);
}

export const duplicateDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/duplicates/" + id;
    return await axios.post(url);
}

export const sendDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications";
    return await axios.post(url, payload);
}

export const updateDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.put(url, payload);
}

export const createDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.post(url, payload);
}

export const getTeams = async (tenantId?:any): Promise<any> => {
    let url = baseAxiosUrl + "/teamdata?tenantId="+tenantId;
    return await axios.get(url);
}

export const getConsentSummaries = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/consentSummaries/" + id;
    return await axios.get(url);
}

export const sendPreview = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/previews";
    // console.log("in api preview",url)
    return await axios.post(url, payload);
}

export const getAuthenticationConsentMetadata = async (windowLocationOriginDomain: string, login_hint: string): Promise<any> => {
    let url = `${baseAxiosUrl}/authenticationMetadata/consentUrl?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${login_hint}`;
    return await axios.get(url, undefined, false);
}

export const sendPdfFile = async (payload: {}): Promise<any> => {
     let url = baseAxiosUrl + "/fileupload/savepdffile";
        console.log("in api survey",url)
    return await axios.post(url, payload,);
}

export const surveyexport = async (id: any): Promise<any> => {
    let url = baseAxiosUrl + "/surveyexport/exportdata?id="+id;
     console.log("in api survey",url,id)
    return await axios.post(url);
}

export const reactionexport = async (id: any): Promise<any> => {
    let url = baseAxiosUrl + "/reactionexport/exportReactionDetail?id="+id;
    console.log("in api reaction",url,id)
    return await axios.post(url);
}

export const recallDataExport = async (id: any): Promise<any> => {
    let url = baseAxiosUrl + "/recallSentNotifications/download/"+id;
    // console.log("in api reaction",url,id)
    return await axios.get(url);
}

export const getSisterTenant = async (): Promise<any> => {
    let url = baseAxiosUrl + "/companycommunicator/tenantlist";
    // console.log("in api",url)
    return await axios.get(url);
}

export const getRecallSentNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/recallSentNotifications";
    return await axios.get(url);
}

export const recalledNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/recallSentNotifications";
    return await axios.post(url, payload);
}

export const getAllTenant = async (): Promise<any> => {
    let url = baseAxiosUrl + "/companycommunicator/allTenantList";
    // console.log("in api",url)
    return await axios.get(url);
}

export const getIPAddress = async (): Promise<any> => {
    let url = "https://ipapi.co/json/";
    //"https://geolocation-db.com/json/";
    // console.log("in api",url)
    return await axios.get(url);
}