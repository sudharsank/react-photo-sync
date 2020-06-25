declare interface IPhotoSyncWebPartStrings {
    PropertyPaneDescription: string;
    BasicGroupName: string;
    ListCreationText: string;
    PropTemplateLibLabel: string;
    PropAzFuncLabel: string;
    PropAzFuncDesc: string;
    PropUseCertLabel: string;
    PropUseCertCallout: string;
    PropDateFormatLabel: string;
    PropInfoDateFormat: string;
    PropInfoUseCert: string;
    PropInfoTemplateLib: string;
    PropInfoNormalUser: string;
    PropAllowedUserInfo: string;

    DefaultAppTitle: string;
    PlaceholderIconText: string;
    PlaceholderDescription: string;
    PlaceholderButtonLabel: string;
    AccessCheckDesc: string;
    SitePrivilegeCheckLabel: string;

    PropsLoader: string;
    PropsUpdateLoader: string;
    AdminConfigHelp: string;
}

declare module 'PhotoSyncWebPartStrings' {
    const strings: IPhotoSyncWebPartStrings;
    export = strings;
}
