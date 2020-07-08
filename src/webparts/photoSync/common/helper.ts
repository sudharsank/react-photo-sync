import { MSGraphClient } from '@microsoft/sp-http';
import { graph } from "@pnp/graph";
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
import { IUserInfo } from './IModel';

const storage = new PnPClientStorage();

const map: any = require('lodash/map');
const intersection: any = require('lodash/intersection');
const orderBy: any = require('lodash/orderBy');

const userDefStorageKey: string = 'userDefaultInfo';
const userCusStorageKey: string = 'userCustomInfo';

export interface IHelper {
    getCurrentUserDefaultInfo: () => Promise<ISiteUserInfo>;
    getCurrentUserCustomInfo: () => Promise<IUserInfo>;
    checkCurrentUserGroup: (allowedGroups: string[], userGroups: string[]) => boolean;
}

export default class Helper implements IHelper {
    private _web: IWeb = null;
    private _graphClient: MSGraphClient = null;

    constructor(weburl?: string, graphClient?: MSGraphClient) {
        this._graphClient = graphClient ? graphClient : null;
        this._web = weburl ? Web(weburl) : sp.web;
    }

    private dataURItoBlob(dataURI) {
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
}