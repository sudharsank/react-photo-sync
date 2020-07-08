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
    PropEnableBUCallout: string;

    DefaultAppTitle: string;
    PlaceholderIconText: string;
    PlaceholderDescription: string;
    PlaceholderButtonLabel: string;
    AccessCheckDesc: string;
    SitePrivilegeCheckLabel: string;

    PPLPickerTitleText: string;
    PropsLoader: string;
    PropsUpdateLoader: string;
    AdminConfigHelp: string;
    AccessDenied: string;

    TabMenu1: string;
    TabMenu2: string;
    TabMenu3: string;
    TabMenu4: string;
    TabMenu5: string;
}

declare module 'PhotoSyncWebPartStrings' {
    const strings: IPhotoSyncWebPartStrings;
    export = strings;
}
