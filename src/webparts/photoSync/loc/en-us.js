define([], function() {
  return {
    PropertyPaneDescription: "",
    BasicGroupName: "Configurations",
    ListCreationText: "Verifying the required list...",
    PropTemplateLibLabel: "Select a library to store the templates",
    PropAzFuncLabel: "Azure Function URL",
    PropAzFuncDesc: "Azure powershell function to update the user profile properties in SharePoint UPS",
    PropUseCertLabel: "Use Certificate for Azure Function authentication",
    PropUseCertCallout: "Turn on this option to use certificate for authenticating SharePoint communication via Azure Function",
    PropDateFormatLabel: "Date format",
    PropInfoDateFormat: "The date format use <strong>momentjs</strong> date format. Please <a href='https://momentjs.com/docs/#/displaying/format/' target='_blank'>click here</a> to get more info on how to define the format. By default the format is '<strong>DD, MMM YYYY hh:mm A</strong>'",
    PropInfoUseCert: "Please <a href='https://www.youtube.com/watch?v=plS_1BsQAto&list=PL-KKED6SsFo8TxDgQmvMO308p51AO1zln&index=2&t=0s' target='_blank'>click here</a> to see how to create Azure powershell function with different authentication mechanism.",
    PropInfoTemplateLib: "Document library to maintain the templates and batch files uploaded. </br>'<strong>SyncJobTemplate</strong>' folder will be created to maintain the templates.</br>'<strong>UPSDataToProcess</strong>' folder will be created to maintain the files uploaded for bulk sync.",
    PropInfoNormalUser: "Sorry, the configuration is enabled only for the site administrators, please contact your site administrator!",
    PropAllowedUserInfo: "Only SharePoint groups are allowed in this setting. Only memebers of the SharePoint groups defined above will have access to this web part.",    
    PropEnableBUCallout: "Turn on to enable bulk update",

    DefaultAppTitle: "SharePoint Profile Photo Sync",
    PlaceholderIconText: "Configure the settings",
    PlaceholderDescription: "Use the configuration settings to map the document library, azure function and other settings.",
    PlaceholderButtonLabel: "Configure",
    AccessCheckDesc: "Checking for access...",
    SitePrivilegeCheckLabel: "Checking site admin privilege...",

    PPLPickerTitleText: "Select users to sync their photos!",
    PropsLoader: "Please wait...",
    PropsUpdateLoader: "Please wait, initializing the job to update the properties",
    AdminConfigHelp: "Please contact your site administrator to configure the webpart.",
    AccessDenied: "Access denied. Please contact your administrator.",

    TabMenu1: "User Selection Photo Sync",
    TabMenu2: "Bulk Photo Sync",
    TabMenu3: "Bulk Files Uploaded",
    // TabMenu4: "Templates Generated",
    TabMenu5: "Sync Status"
  }
});