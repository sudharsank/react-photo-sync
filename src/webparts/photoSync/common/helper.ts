import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { MSGraphClient } from '@microsoft/sp-http';
import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/graph/groups";
import { sp } from '@pnp/sp';
import "@pnp/sp/profiles";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { Web, IWeb } from "@pnp/sp/webs";
import { ISiteUserInfo, ISiteUser } from "@pnp/sp/site-users/types";
import { PnPClientStorage, dateAdd } from '@pnp/common';
import { IUserInfo, IUserPickerInfo, SyncType, JobStatus, IAzFuncValues } from './IModel';
import * as moment from 'moment';

import "@pnp/sp/search";
import { SearchQueryBuilder, SearchResults, ISearchQuery } from "@pnp/sp/search";

const storage = new PnPClientStorage();

const map: any = require('lodash/map');
const intersection: any = require('lodash/intersection');
const orderBy: any = require('lodash/orderBy');
const chunk: any = require('lodash/chunk');
const batchItemLimit: number = 18;
const userBatchLimit: number = 6;

const userDefStorageKey: string = 'userDefaultInfo';
const userCusStorageKey: string = 'userCustomInfo';

export interface IHelper {
    getLibraryDetails: (listid: string) => Promise<any>;
    dataURItoBlob: (dataURI: any) => Blob;
    getCurrentUserDefaultInfo: () => Promise<ISiteUserInfo>;
    getCurrentUserCustomInfo: () => Promise<IUserInfo>;
    checkCurrentUserGroup: (allowedGroups: string[], userGroups: string[]) => boolean;
    getUsersInfo: (UserIds: string[]) => Promise<any[]>;
    getUserPhotoFromAADForDisplay: (users: IUserPickerInfo[]) => Promise<any[]>;
    getAndStoreUserThumbnailPhotos: (users: IUserPickerInfo[], tempLibId: string) => Promise<IAzFuncValues[]>;
    createSyncItem: (syncType: SyncType) => Promise<number>;
    updateSyncItem: (itemid: number, inputJson: string) => void;
    getAllJobs: () => Promise<any[]>;
    runAzFunction: (httpClient: HttpClient, inputData: any, azFuncUrl: string, itemid: number) => void;
}

export default class Helper implements IHelper {
    private _web: IWeb = null;
    private _graphClient: MSGraphClient = null;
    private _graphUrl: string = "https://graph.microsoft.com/v1.0";
    private web_ServerRelativeURL: string = '';
    private TPhotoFolderName: string = 'UserPhotos';
    private Lst_SyncJobs = 'UPS Photo Sync Jobs';

    constructor(webRelativeUrl: string, weburl?: string, graphClient?: MSGraphClient) {
        this._graphClient = graphClient ? graphClient : null;
        this._web = weburl ? Web(weburl) : sp.web;
        this.web_ServerRelativeURL = webRelativeUrl;
    }

    private _demo = async () => {
        const results2: SearchResults = await sp.search(<ISearchQuery>{
            Querytext: "mani*",
            RowLimit: 10,
            EnableInterleaving: true,
            SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
            SelectProperties: ['PreferredName', 'AccountName', 'PictureURL', 'PictureHeight', 'PictureThumnailURL', 'PictureWidth', 'Size', 'DisplayDate']
        });
        console.log(results2.PrimarySearchResults);
    }

    /**
     * Get temp library details
     * @param listid Temporary library
     */
    public getLibraryDetails = async (listid: string): Promise<string> => {
        let retFolderPath: string = '';
        let listDetails = await this._web.lists.getById(listid).get();
        retFolderPath = listDetails.DocumentTemplateUrl.replace('/Forms/template.dotx', '') + '/' + this.TPhotoFolderName;
        return retFolderPath;
    }

    public dataURItoBlob = (dataURI): Blob => {
        // convert base64/URLEncoded data component to raw binary data held in a string
        var byteString;
        if (dataURI.split(',')[0].indexOf('base64') >= 0)
            byteString = atob(dataURI.split(',')[1]);
        else
            byteString = unescape(dataURI.split(',')[1]);
        // separate out the mime component
        var mimeString = dataURI.split(',')[0].split(':')[1].split(';')[0];
        // write the bytes of the string to a typed array
        var ia = new Uint8Array(byteString.length);
        for (var i = 0; i < byteString.length; i++) {
            ia[i] = byteString.charCodeAt(i);
        }
        return new Blob([ia], { type: mimeString });
    }
    /**
     * Get current logged in user default info.
     */
    public getCurrentUserDefaultInfo = async (): Promise<ISiteUserInfo> => {
        //return await this._web.currentUser.get();
        let currentUserInfo: ISiteUserInfo = storage.local.get(userDefStorageKey);
        if (!currentUserInfo) {
            currentUserInfo = await this._web.currentUser.get();
            storage.local.put(userDefStorageKey, currentUserInfo, dateAdd(new Date(), 'hour', 1));
        }
        return currentUserInfo;
    }
    /**
     * Get current logged in user custom information.
     */
    public getCurrentUserCustomInfo = async (): Promise<IUserInfo> => {
        let currentUserInfo = await this._web.currentUser.get();
        let currentUserGroups = await this._web.currentUser.groups.get();
        return ({
            ID: currentUserInfo.Id,
            Email: currentUserInfo.Email,
            LoginName: currentUserInfo.LoginName,
            DisplayName: currentUserInfo.Title,
            IsSiteAdmin: currentUserInfo.IsSiteAdmin,
            Groups: map(currentUserGroups, 'LoginName'),
            Picture: '/_layouts/15/userphoto.aspx?size=S&username=' + currentUserInfo.UserPrincipalName,
        });
    }
    /**
     * Check current user is a member of groups or not.
     */
    public checkCurrentUserGroup = (allowedGroups: string[], userGroups: string[]): boolean => {
        if (userGroups.length > 0) {
            let diff: string[] = intersection(allowedGroups, userGroups);
            if (diff && diff.length > 0) return true;
        }
        return false;
    }
    /**
     * Get user profile photos from Azure AD
     */
    public getUserPhotoFromAADForDisplay = async (users: IUserPickerInfo[]): Promise<any[]> => {
        return new Promise(async (res, rej) => {
            if (users && users.length > 0) {
                let requests: any[] = [];
                let finalResponse: any[] = [];
                if (users.length > batchItemLimit) {
                    let chunkUserArr: any[] = chunk(users, batchItemLimit);
                    Promise.all(chunkUserArr.map(async chnkdata => {
                        requests = [];
                        chnkdata.map((user: IUserPickerInfo) => {
                            let upn: string = user.LoginName.split('|')[2];
                            requests.push({
                                id: `${user.LoginName}`,
                                method: 'GET',
                                responseType: 'blob',
                                headers: { "Content-Type": "image/jpeg" },
                                url: `/users/${upn}/photos/$value`
                            });
                        });
                        let photoReq: any = { requests: requests };
                        let graphRes: any = await this._graphClient.api('$batch').post(photoReq);
                        finalResponse.push(graphRes);
                    })).then(() => {
                        res(finalResponse);
                    });
                } else {
                    users.map((user: IUserPickerInfo) => {
                        let upn: string = user.LoginName.split('|')[2];
                        requests.push({
                            id: `${user.LoginName}`,
                            method: 'GET',
                            responseType: 'blob',
                            headers: { "Content-Type": "image/jpeg" },
                            url: `/users/${upn}/photo/$value`
                        });
                    });
                    let photoReq: any = { requests: requests };
                    finalResponse.push(await this._graphClient.api('$batch').post(photoReq));
                    res(finalResponse);
                }
            }
        });
    }
    /**
     * Get user info based on UserID
     */
    public getUsersInfo = async (userids: string[]): Promise<any[]> => {
        return new Promise(async (res, rej) => {
            let finalResponse: any[] = [];
            if (userids.length > batchItemLimit) {

            } else {
                let batch = sp.createBatch();
                userids.map((userid: string) => {
                    sp.web.siteUsers.getByLoginName(`i:0#.f|membership|${userid}`).inBatch(batch).get().then((userinfo) => {
                        console.log(userinfo);
                        if (userinfo && userinfo.Title) {
                            finalResponse.push({
                                'loginname': userid,
                                'title': userinfo.Title,
                                'status': 'Success'
                            });
                        }
                    }).catch((e) => {
                        finalResponse.push({
                            'loginname': userid,
                            'title': 'User not found!',
                            'status': 'Failed'
                        });
                    });
                });
                batch.execute().then(() => {
                    res(finalResponse);
                }).catch(() => {
                    res(finalResponse);
                });
            }
        });
    }
    /**
     * Get thumbnail photos for the users.
     * @param users List of users
     */
    public getAndStoreUserThumbnailPhotos = async (users: IUserPickerInfo[], tempLibId: string): Promise<IAzFuncValues[]> => {
        let retVals: IAzFuncValues[] = [];
        return new Promise(async (res, rej) => {
            let tempLibUrl: string = await this.getLibraryDetails(tempLibId);
            if (users && users.length > 0) {
                let requests: any[] = [];
                let finalResponse: any[] = [];
                if (users.length > userBatchLimit) {
                    let chunkUserArr: any[] = chunk(users, userBatchLimit);
                    Promise.all(chunkUserArr.map(async chnkdata => {
                        requests = [];
                        chnkdata.map((user: IUserPickerInfo) => {
                            let upn: string = user.LoginName.split('|')[2];
                            requests.push({
                                id: `${user.LoginName}_1`,
                                method: 'GET',
                                responseType: 'blob',
                                headers: { "Content-Type": "image/jpeg" },
                                url: `/users/${upn}/photos/48x48/$value`
                            }, {
                                id: `${user.LoginName}_2`,
                                method: 'GET',
                                responseType: 'blob',
                                headers: { "Content-Type": "image/jpeg" },
                                url: `/users/${upn}/photos/96x96/$value`
                            }, {
                                id: `${user.LoginName}_3`,
                                method: 'GET',
                                responseType: 'blob',
                                headers: { "Content-Type": "image/jpeg" },
                                url: `/users/${upn}/photos/240x240/$value`
                            });
                        });
                        let photoReq: any = { requests: requests };
                        let graphRes: any = await this._graphClient.api('$batch').post(photoReq);
                        finalResponse.push(graphRes);
                    })).then(async () => {
                        retVals = await this.saveThumbnailPhotosInDocLib(finalResponse, tempLibUrl);
                    });
                } else {
                    users.map((user: IUserPickerInfo) => {
                        let upn: string = user.LoginName.split('|')[2];
                        requests.push({
                            id: `${user.LoginName}_1`,
                            method: 'GET',
                            responseType: 'blob',
                            headers: { "Content-Type": "image/jpeg" },
                            url: `/users/${upn}/photos/48x48/$value`
                        }, {
                            id: `${user.LoginName}_2`,
                            method: 'GET',
                            responseType: 'blob',
                            headers: { "Content-Type": "image/jpeg" },
                            url: `/users/${upn}/photos/96x96/$value`
                        }, {
                            id: `${user.LoginName}_3`,
                            method: 'GET',
                            responseType: 'blob',
                            headers: { "Content-Type": "image/jpeg" },
                            url: `/users/${upn}/photos/240x240/$value`
                        });
                    });
                    let photoReq: any = { requests: requests };
                    finalResponse.push(await this._graphClient.api('$batch').post(photoReq));
                    retVals = await this.saveThumbnailPhotosInDocLib(finalResponse, tempLibUrl);
                }
            }
            res(retVals);
        });
    }
    /**
     * Add thumbnails to the configured document library
     */
    private saveThumbnailPhotosInDocLib = async (thumbnails: any[], tempLibName: string): Promise<IAzFuncValues[]> => {
        let retVals: IAzFuncValues[] = [];
        if (thumbnails && thumbnails.length > 0) {
            thumbnails.map(res => {
                if (res.responses && res.responses.length > 0) {
                    res.responses.map(async thumbnail => {
                        if (!thumbnail.body.error) {
                            let username: string = thumbnail.id.split('_')[0].split('|')[2];
                            let userFilename: string = username.replace(/[@.]/g, '_');
                            let filecontent = this.dataURItoBlob("data:image/jpg;base64," + thumbnail.body);
                            let partFileName = '';
                            retVals.push({
                                userid: username,
                                picturename: userFilename
                            });
                            if (thumbnail.id.indexOf('_1') > 0) partFileName = 'SThumb.jpg';
                            else if (thumbnail.id.indexOf('_2') > 0) partFileName = "MThumb.jpg";
                            else if (thumbnail.id.indexOf('_3') > 0) partFileName = "LThumb.jpg";
                            await sp.web.getFolderByServerRelativeUrl(decodeURI(`${tempLibName}/`))
                                .files
                                .add(decodeURI(`${tempLibName}/${userFilename}_` + partFileName), filecontent, true);
                        }
                    });
                }
            });
        }
        return retVals;
    }
    /**
     * Create a sync item
     */
    public createSyncItem = async (syncType: SyncType): Promise<number> => {
        let returnVal: number = 0;
        let itemAdded = await this._web.lists.getByTitle(this.Lst_SyncJobs).items.add({
            Title: `SyncJob_${moment().format("MMDDYYYYhhmm")}`,
            Status: JobStatus.Submitted.toString(),
            SyncType: syncType.toString()
        });
        returnVal = itemAdded.data.Id;
        return returnVal;
    }
    /**
     * Update Sync item with the input data to sync
     */
    public updateSyncItem = async (itemid: number, inputJson: string) => {
        await this._web.lists.getByTitle(this.Lst_SyncJobs).items.getById(itemid).update({
            SyncData: inputJson
        });
    }
    /**
     * Update Sync item with the error status
     */
    public updateSyncItemStatus = async (itemid: number, errMsg: string) => {
        await this._web.lists.getByTitle(this.Lst_SyncJobs).items.getById(itemid).update({
            Status: JobStatus.Error,
            ErrorMessage: errMsg
        });
    }
    /**
     * Get all the jobs items
     */
    public getAllJobs = async (): Promise<any[]> => {
        return await this._web.lists.getByTitle(this.Lst_SyncJobs).items
            .select('ID', 'Title', 'SyncedData', 'Status', 'ErrorMessage', 'SyncType', 'Created', 'Author/Title', 'Author/Id', 'Author/EMail')
            .expand('Author')
            .getAll();
    }
    /**
     * Azure function to update the UPS Photo properties.
     */
    public runAzFunction = async (httpClient: HttpClient, inputData: any, azFuncUrl: string, itemid: number) => {
        const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-type", "application/json");
        requestHeaders.append("Cache-Control", "no-cache");
        const postOptions: IHttpClientOptions = {
            headers: requestHeaders,
            body: `${inputData}`
        };
        let response: HttpClientResponse = await httpClient.post(azFuncUrl, HttpClient.configurations.v1, postOptions);
        if (!response.ok) {
            await this.updateSyncItemStatus(itemid, `${response.status} - ${response.statusText}`);
        }
        console.log("Azure Function executed");
    }

    // public async componentDidMount() {
    //     // let currentUser = await graph.me.get();
    //     // console.log(currentUser);
    //     // let currentUserPhoto = await currentUser.photo;
    //     // console.log(currentUserPhoto);
    //     // let someUser = await graph.users.getById("adelev@o365practice.onmicrosoft.com").get();
    //     // console.log(someUser);
    //     // let someUserPhoto = await someUser.photos;
    //     // console.log(someUserPhoto);
    //     let res = await this.props.client.api('https://graph.microsoft.com/v1.0/me/photos/240x240/$value').responseType('string').get();
    //     console.log(res);
    //     // this._setPhoto(res);
    //     let photoTReq: any = {
    //         requests: [{
    //             id: '1',
    //             method: 'GET',
    //             responseType: 'blob',
    //             "headers": {
    //                 "Content-Type": "image/jpeg"
    //             },
    //             url: '/me/photos/240x240/$value'
    //         }, {
    //             id: '2',
    //             method: 'GET',
    //             "headers": {
    //                 "Content-Type": "image/jpeg"
    //             },
    //             url: '/me/photos/96x96/$value'
    //         }]
    //     };
    //     let res1 = await this.props.client.api('$batch').post(photoTReq);
    //     console.log(res1);
    //     var blob = new Blob()
    //     res1.responses.map(res => {
    //         let filecontent = this.dataURItoBlob("data:image/jpg;base64," + res.body);
    //         let partFileName = '';
    //         if (res.id == "1") partFileName = 'LThumb.jpg';
    //         else if (res.id == "2") partFileName = "MThumb.jpg";
    //         else if (res.id == "3") partFileName = "SThumb.jpg";
    //         sp.web.getFolderByServerRelativeUrl(decodeURI('/sites/ModernDev/Sample%20Documents/UserPhotos/'))
    //             .files
    //             .add(decodeURI('/sites/ModernDev/Sample%20Documents/UserPhotos/revathy_o365practice_onmicrosoft_com_' + partFileName), filecontent, true);
    //     });
    // }

}