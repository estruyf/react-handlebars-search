import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
import { ISearchVisualizerWebPartProps } from "../ISearchVisualizerWebPartProps";

export interface ISearchVisualizerProps extends ISearchVisualizerWebPartProps {
  audienceTargeting: string;
  audienceTargetingAll: string;
  audienceTargetingBooleanOperator: string;
  context: WebPartContext;

  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
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


