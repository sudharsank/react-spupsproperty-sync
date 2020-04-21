define([], function() {
  return {
    PropertyPaneDescription: "",
    BasicGroupName: "Configurations",
    ListCreationText: "Verifying the required list and loading the properties...",
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
    
    PlaceholderIconText: "Configure the settings",
    PlaceholderDescription: "Use the configuration settings to map the document library required to store the property mapping templates.",
    PlaceholderButtonLabel: "Configure",
    DefaultAppTitle: "SharePoint Profile Property Sync",
    JobResultsDialogTitle: "Users list with properties updated!",
    JobsListSearchPH: "Search by Title, SyncType, Author, Status...",
    TemplateListSearchPH: "Search by Name, Author...",
    TemplateStructureDialogTitle: "Properties defined in the template!",
    BulkSyncDataDialogTitle: "Data defined in the file!",

    GenerateTemplateLoader: "Wait, generating the template...",
    UploadDataToSyncLoader: "Wait, uploading data for syncing",
    PropsLoader: "Please wait...",
    PropsUpdateLoader: "Please wait, initializing the job to update the properties",
    JobsListLoaderDesc: "Loading the jobs list...",
    JobResultsLoaderDesc: "Loading the results...",
    TemplateListLoaderDesc: "Loading the templates...",
    TemplatePropsLoaderDesc: "Loading properties, please wait...",
    BulkSyncListLoaderDesc: "Loading the bulk sync files...",
    BulkSyncFileDataLoaderDesc: "Loading data, please wait...",
    AccessCheckDesc: "Checking for access...",
    SitePrivilegeCheckLabel: "Checking site admin privilege...",

    BtnGenerateJSON: "Generate JSON",
    BtnGenerateCSV: "Generate CSV",
    BtnSaveForManual: "Save for Manual Update",
    BtnPropertyMapping: "Property Mapping",
    BtnUploadDataForSync: "Upload Data to Sync",
    BtnUpdateUserProps: "Update User Properties",
    BtnManualProps: "Initialize Manual Properties",
    BtnAzureProps: "Get Azure Properties",

    PnlHeaderText: "Property Mappings",
    TblColHeadAzProperty: "Azure Property",
    TblColHeadSPProperty: "SharePoint Property",
    TblColHeadEnableDisable: "Enabled/Disabled",
    PPLPickerTitleText: "Select users to update their properties",

    EmptyPropertyMappings: "No active property mappings found. Please navigate to 'Sync Properties Mapping' list or contact your administrator to activate the properties.",
    TemplateDownloaded: "Please use the downloaded file to update the User properties!",
    EmptyDataText: "Empty!",
    EmptyDataWarningMsg: "Columns with empty values are not considered for update!",
    EmptyTable: "Sorry, no data to be displayed!",
    EmptyFile: "Oops, the file is empty",
    EmptySearchResults: "Sorry, no data found. Displaying all the data",
    UserListChanges: "Changes in user list, please remove the user from the table manually or reinitialize or get the Azure properties again!",
    UserListEmpty: "Since all the users have been removed, the table has been cleared!",
    JobIntializedSuccess: "Property sync job has been initialized. Track the status of the job on the 'Sync Jobs' tab!",
    AdminConfigHelp: "Please contact your site administrator to configure the webpart.",
    AccessDenied: "Access denied. Please contact your administrator.",
    SyncFailedErrorMessage: "Oops, there is an error while updating the properties. Error Message:",

    TabMenu1: "Manual or Azure Property Sync",
    TabMenu2: "Bulk Sync",
    TabMenu3: "Bulk Files Uploaded",
    TabMenu4: "Templates Generated",
    TabMenu5: "Sync Status"
  }
});