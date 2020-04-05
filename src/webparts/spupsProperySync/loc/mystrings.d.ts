declare interface ISpupsProperySyncWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  GenerateTemplateLoader: string;
  UploadDataToSyncLoader: string;
  PropsLoader: string;

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
  UserListChanges: string;
  UserListEmpty: string;
}

declare module 'SpupsProperySyncWebPartStrings' {
  const strings: ISpupsProperySyncWebPartStrings;
  export = strings;
}
