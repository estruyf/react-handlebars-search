import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISearchVisualizerProps {
    title: string;
    query: string;
    maxResults: number;
    sorting: string;
    debug: boolean;
    external: string;
    scriptloading: boolean;
    duplicates: boolean;
    privateGroups: boolean;
    audienceTargeting: string;
    audienceTargetingAll: string;
    audienceTargetingBooleanOperator: string;
    context: WebPartContext;
}

export interface ISearchVisualizerState {
    loading?: boolean;
    template?: string;
    error?: string;
    showError?: boolean;
    showScriptDialog?: boolean;
}

export interface IMetadata {
    fields: string[];
    resources: ITemplateResource[];
}

export interface ITemplateResource {
    key: string;
    values: ILocaleResource;
}

export interface ILocaleResource {
    [locale: string]: string;
}

export interface ISPUser {
    username?: string;
    displayName?: string;
    email?: string;
}

export interface ISPUrl {
    url?: string;
    description?: string;
}


