define([], function () {
    return {
        'PropertyPaneDescription': 'Search Web Part Settings',
        'QueryGroupName': 'Search query settings',
        'AudienceGroupName': 'Audience Targeting settings',
        'TemplateGroupName': 'Template settings',
        'TitleFieldLabel': 'Web part title',
        'QueryFieldLabel': 'Specify your search query',
        'QueryFieldDescription': 'You can make use of following tokens: {Site} - {SiteCollection} - {Today} or {Today+NR} or {Today-NR} - {CurrentDisplayLanguage} - {User}, {User.Name}, {User.Email}',
        'FieldsMaxResults': 'Number of results to render per page',
        'SortingFieldLabel': 'Sorting (MP:ascending or descending) - example: lastmodifiedtime:ascending,author:descending',
        'DebugFieldLabel': 'Show debug output?',
        'DebugFieldLabelOn': 'Yes',
        'DebugFieldLabelOff': 'No',
        'ExternalFieldLabel': 'Specify the external template URL',
        'ScriptloadingFieldLabel': 'Enable script loading from the template?',
        'ScriptloadingFieldLabelOn': 'Danger mode',
        'ScriptloadingFieldLabelOff': 'Safe mode',
        'DuplicatesFieldLabel': 'Trim duplicate results?',
        'DuplicatesFieldLabelOn': 'Yes',
        'DuplicatesFieldLabelOff': 'No',
        'PrivateGroupsFieldLabel': 'Search through content of private groups?',
        'PrivateGroupsFieldLabelOn': 'Yes',
        'PrivateGroupsFieldLabelOff': 'No, only public groups',
        'AudienceColumnMappingLabel': 'Audience Managed Property to User Profile Property Mapping',
        'AudienceColumnMappingDescription': 'On each line, map each audience managed property to the corresponding user profile property. Example: {"ManagedPropertyAlias":"UserProfileProperty"}',
        'AudienceAllValueLabel': 'Audience managed property value to indicate content targeted to everyone',
        'AudienceAllValueDescription': 'Enter the managed property and its value that will indicate which content is targeted to everyone. Example: {"ManagedPropertyAlias":"All"}',
        'AudienceBooleanOperatorLabel': 'Select the boolean operator for the above mappings',

        'QuertValidationEmpty': 'Please specify a search query',
        'TemplateValidationEmpty': 'Please specify the URL of your template',
        'TemplateValidationHTML': 'Please use an HTML file for the template',

        'ScriptsDialogHeader': 'Script tags found in your template',
        'ScriptsDialogSubText': 'If you want to be able to run the scripts from your template, you have to activate the script loading option.'
    }
});
