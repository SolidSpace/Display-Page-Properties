declare interface IDisplayPagePropertyWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  CodeEditorButtonLabel: string;
  CodeEditorPanelTitle: string;
  DescriptionFieldLabel: string;
  DefaultTemplateHeadline:string;
  TemplateTextStrings: any;
  ConfigurePageProperties:string;
  LayoutTemplateLabel: string;
  LayoutTemplateLabelCustom: string;
  LayoutTemplateLabelDebug: string;
  LayoutTemplateLabelRows: string;
  LayoutTemplateLabelTable: string;
  DisplayPagePropertyStrings:any;
  PlaceholderButtonLabel:string;
}

declare module 'DisplayPagePropertyWebPartStrings' {
  const strings: IDisplayPagePropertyWebPartStrings;
  export = strings;
}
