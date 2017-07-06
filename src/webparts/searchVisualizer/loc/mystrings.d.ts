declare interface ISearchVisualizerStrings {
  PropertyPaneDescription: string;
  QueryGroupName: string;
  TemplateGroupName: string;
  TitleFieldLabel: string;
  QueryFieldLabel: string;
  QueryFieldDescription: string;
  FieldsMaxResults: string;
  SortingFieldLabel: string;
  DebugFieldLabel: string;
  ExternalFieldLabel: string;
  ScriptloadingFieldLabel: string;
  QuertValidationEmpty: string;
  TemplateValidationEmpty: string;
  TemplateValidationHTML: string;
  ScriptsDialogHeader: string;
  ScriptsDialogSubText: string;
}

declare module 'searchVisualizerStrings' {
  const strings: ISearchVisualizerStrings;
  export = strings;
}
