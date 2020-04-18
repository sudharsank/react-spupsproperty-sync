declare interface ISpupsProperySyncWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PropTemplateLibLabel: string;
  PropAzFuncLabel: string;
  PropAzFuncDesc: string;
  PropUseCertLabel: string;
  PropUseCertCallout: string;
  PropDateFormatLabel: string;
  PropInfoDateFormat: string;
  PropInfoUseCert: string;
  PropInfoTemplateLib: string;

  PlaceholderIconText: string;
  PlaceholderDescription: string;
  PlaceholderButtonLabel: string;
  DefaultAppTitle: string;
  JobResultsDialogTitle: string;
  JobsListSearchPH: string;
  TemplateListSearchPH: string;
  TemplateStructureDialogTitle: string;
  BulkSyncDataDialogTitle: string;
  BulkSyncFileDataLoaderDesc: string;

  GenerateTemplateLoader: string;
  UploadDataToSyncLoader: string;
  PropsLoader: string;
  PropsUpdateLoader: string;
  JobsListLoaderDesc: string;
  JobResultsLoaderDesc: string;
  TemplateListLoaderDesc: string;  
  TemplatePropsLoaderDesc: string;
  BulkSyncListLoaderDesc: string;

  BtnGenerateJSON: string;
  BtnGenerateCSV: string;
  BtnSaveForManual: string;
  BtnPropertyMapping: string;
  BtnUploadDataForSync: string;
  BtnUpdateUserProps: string;
  BtnManualProps: string;
  BtnAzureProps: string;

  PnlHeaderText: string;
  TblColHeadAzProperty: string;
  TblColHeadSPProperty: string;
  TblColHeadEnableDisable: string;
  PPLPickerTitleText: string;

  EmptyPropertyMappings: string;
  TemplateDownloaded: string;
  EmptyDataText: string;
  EmptyDataWarningMsg: string;
  EmptyTable: string;
  EmptyFile: string;
  EmptySearchResults: string;
  UserListChanges: string;
  UserListEmpty: string;
  JobIntializedSuccess: string;  
}

declare module 'SpupsProperySyncWebPartStrings' {
  const strings: ISpupsProperySyncWebPartStrings;
  export = strings;
}
