import { MSGraphClient } from '@microsoft/sp-http';
import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/graph/groups";
import { sp } from '@pnp/sp';
import "@pnp/sp/profiles";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { Web, IWeb } from "@pnp/sp/webs";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { PnPClientStorage, dateAdd } from '@pnp/common';
import { IUserInfo, IUserPickerInfo } from './IModel';

import "@pnp/sp/search";
import { SearchQueryBuilder, SearchResults, ISearchQuery } from "@pnp/sp/search";

const storage = new PnPClientStorage();

const map: any = require('lodash/map');
const intersection: any = require('lodash/intersection');
const orderBy: any = require('lodash/orderBy');
const chunk: any = require('lodash/chunk');

const userDefStorageKey: string = 'userDefaultInfo';
const userCusStorageKey: string = 'userCustomInfo';

export interface IHelper {
    dataURItoBlob: (dataURI: any) => Blob;
    getCurrentUserDefaultInfo: () => Promise<ISiteUserInfo>;
    getCurrentUserCustomInfo: () => Promise<IUserInfo>;
    checkCurrentUserGroup: (allowedGroups: string[], userGroups: string[]) => boolean;
    getUserPhotoFromAADForDisplay: (users: IUserPickerInfo[]) => Promise<any[]>;
}

export default class Helper implements IHelper {
    private _web: IWeb = null;
    private _graphClient: MSGraphClient = null;
    private _graphUrl: string = "https://graph.microsoft.com/v1.0";

    constructor(weburl?: string, graphClient?: MSGraphClient) {
        this._graphClient = graphClient ? graphClient : null;
        this._web = weburl ? Web(weburl) : sp.web;
        this._demo();
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
        let batchItems: number = 15;
        return new Promise(async (res, rej) => {
            if (users && users.length > 0) {
                let requests: any[] = [];
                let finalResponse: any[] = [];
                if (users.length > batchItems) {
                    let chunkUserArr: any[] = chunk(users, batchItems);
                    Promise.all(chunkUserArr.map(async chnkdata => {
                        requests = [];
                        chnkdata.map((user: IUserPickerInfo) => {
                            let upn: string = user.LoginName.split('|')[2];
                            requests.push({
                                id: `${user.LoginName}`,
                                method: 'GET',
                                responseType: 'blob',
                                headers: { "Content-Type": "image/jpeg" },
                                url: `/users/${upn}/photos/96x96/$value`
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