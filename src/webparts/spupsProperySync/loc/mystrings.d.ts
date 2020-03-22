declare interface ISpupsProperySyncWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  GenerateTemplateLoader: string;
  UploadDataToSyncLoader: string;

  BtnGenerateJSON: string;
  BtnGenerateCSV: string;
  BtnSaveForManual: string;
  BtnPropertyMapping: string;
  BtnUploadDataForSync: string;

  PnlHeaderText: string;
  TblColHeadAzProperty: string;
  TblColHeadSPProperty: string;
  TblColHeadEnableDisable: string;

  EmptyDataText: string;
  EmptyDataWarningMsg: string;
}

declare module 'SpupsProperySyncWebPartStrings' {
  const strings: ISpupsProperySyncWebPartStrings;
  export = strings;
}
