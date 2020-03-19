declare interface ISpupsProperySyncWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  GenerateTemplateLoader: string;
  BtnGenerateJSON: string;
  BtnGenerateCSV: string;
  BtnPropertyMapping: string;

  PnlHeaderText: string;
  TblColHeadAzProperty: string;
  TblColHeadSPProperty: string;
  TblColHeadEnableDisable: string;
}

declare module 'SpupsProperySyncWebPartStrings' {
  const strings: ISpupsProperySyncWebPartStrings;
  export = strings;
}
