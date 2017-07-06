import { IWebPartContext } from "@microsoft/sp-webpart-base";

export interface ISearchVisualizerProps {
  title: string;
  query: string;
  maxResults: number;
  sorting: string;
  debug: boolean;
  external: string;
  scriptloading: boolean;
  context: IWebPartContext;
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
}
