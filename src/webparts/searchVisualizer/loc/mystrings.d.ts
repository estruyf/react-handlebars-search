declare interface ISearchVisualizerStrings {
    /* Fields */
    PropertyPaneDescription: string;
    QueryGroupName: string;
    AudienceGroupName: string;
    TemplateGroupName: string;
    TitleFieldLabel: string;
    QueryFieldLabel: string;
    QueryFieldDescription: string;
    FieldsMaxResults: string;
    SortingFieldLabel: string;
    DebugFieldLabel: string;
    DebugFieldLabelOn: string;
    DebugFieldLabelOff: string;
    ExternalFieldLabel: string;
    ScriptloadingFieldLabel: string;
    ScriptloadingFieldLabelOn: string;
    ScriptloadingFieldLabelOff: string;
    DuplicatesFieldLabel: string;
    DuplicatesFieldLabelOn: string;
    DuplicatesFieldLabelOff: string;
    PrivateGroupsFieldLabel: string;
    PrivateGroupsFieldLabelOn: string;
    PrivateGroupsFieldLabelOff: string;
    AudienceColumnMappingLabel: string;
    AudienceColumnMappingDescription: string;
    AudienceAllValueLabel: string;
    AudienceAllValueDescription: string;
    AudienceBooleanOperatorLabel: string;

    /* Validation */
    QuertValidationEmpty: string;
    TemplateValidationEmpty: string;
    TemplateValidationHTML: string;

    /* Dialog */
    ScriptsDialogHeader: string;
    ScriptsDialogSubText: string;
}

declare module 'searchVisualizerStrings' {
    const strings: ISearchVisualizerStrings;
    export = strings;
}
