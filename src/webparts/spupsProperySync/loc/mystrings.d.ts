declare interface ISpupsProperySyncWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  GenerateTemplateLoader: string;
  UploadDataToSyncLoader: string;

  BtnGenerateJSON: string;
  BtnGenerateCSV: string;
  BtnPropertyMapping: string;
  BtnUploadDataForSync: string;

  PnlHeaderText: string;
  TblColHeadAzProperty: string;
  TblColHeadSPProperty: string;
  TblColHeadEnableDisable: string;

  EmptyDataText: string;
}

declare module 'SpupsProperySyncWebPartStrings' {
  const strings: ISpupsProperySyncWebPartStrings;
  export = strings;
}
