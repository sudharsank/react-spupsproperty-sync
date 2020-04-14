declare interface ISpupsProperySyncWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;

  PlaceholderIconText: string;
  PlaceholderDescription: string;
  PlaceholderButtonLabel: string;
  DefaultAppTitle: string;
  JobResultsDialogTitle: string;
  JobsListSearchPH: string;

  GenerateTemplateLoader: string;
  UploadDataToSyncLoader: string;
  PropsLoader: string;
  PropsUpdateLoader: string;
  JobsListLoaderDesc: string;
  JobResultsLoaderDesc: string;
  TemplateListLoaderDesc: string;
  TemplateListSearchPH: string;

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

  EmptyDataText: string;
  EmptyDataWarningMsg: string;
  EmptyTable: string;
  EmptyFile: string;
  EmptySearchResults: string;
  UserListChanges: string;
  UserListEmpty: string;
  JobIntializedSuccess: string;

  PropTemplateLibLabel: string;
}

declare module 'SpupsProperySyncWebPartStrings' {
  const strings: ISpupsProperySyncWebPartStrings;
  export = strings;
}
