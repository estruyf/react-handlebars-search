import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'searchVisualizerStrings';
import SearchVisualizer from './components/SearchVisualizer';
import { ISearchVisualizerProps } from './components/ISearchVisualizerProps';
import { ISearchVisualizerWebPartProps } from './ISearchVisualizerWebPartProps';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Text } from '@microsoft/sp-core-library';

export const USERPROFILE_KEY = 'SearchVisualizerWebPart:UserProfileData';

require('./styles/fabric-9.6.1.scoped.css');
export default class SearchVisualizerWebPart extends BaseClientSideWebPart<ISearchVisualizerWebPartProps> {
  private propertyFieldCollectionData = null;
  private customCollectionFieldType = null;

  public render(): void {
    const element: React.ReactElement<ISearchVisualizerProps> = React.createElement(
      SearchVisualizer,
      {
        ...this.properties,
        sorting: this.getSortingOption(),
        audienceTargetingBooleanOperator: this.properties.audienceBooleanOperator ? this.properties.audienceBooleanOperator : 'OR',
        context: this.context
      }
    );
    let domElement: HTMLElement = this.domElement;

    const userProfileData = window.sessionStorage ? sessionStorage.getItem(USERPROFILE_KEY) : null;
    if (this.properties.audienceColumnMapping && this.properties.audienceColumnAllValue && !userProfileData && window.sessionStorage) {
      // get user profile properties if not in session storage and then process search results
      this._getUserProfileProperties().then((result) => {
        if (result.UserProfileProperties) {
          sessionStorage.setItem(USERPROFILE_KEY, JSON.stringify(result.UserProfileProperties));
        }
        ReactDom.render(element, domElement);
      });
    } else {
      ReactDom.render(element, domElement);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Load property pane resources
   */
  protected async loadPropertyPaneResources(): Promise<void> {
    const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import (
      /* webpackChunkName: 'pnp-controls-collectiondata' */
      '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
    );

    this.propertyFieldCollectionData = PropertyFieldCollectionData;
    this.customCollectionFieldType = CustomCollectionFieldType;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.QueryGroupName,
              groupFields: [
                PropertyPaneTextField('query', {
                  label: strings.QueryFieldLabel,
                  description: strings.QueryFieldDescription,
                  multiline: true,
                  onGetErrorMessage: this._queryValidation,
                  deferredValidationTime: 500
                }),
                PropertyPaneSlider('maxResults', {
                  label: strings.FieldsMaxResults,
                  min: 1,
                  max: 50
                }),
                // PropertyPaneTextField('sorting', {
                //   label: strings.SortingFieldLabel
                // }),
                this.propertyFieldCollectionData("mpSorting", {
                  key: "mpSorting",
                  label: strings.SortingPanelLabel,
                  panelHeader: strings.SortingPanelHeader,
                  panelDescription: strings.SortingPanelDescription,
                  manageBtnLabel: strings.ManageSortingBtnLabel,
                  value: this.properties.mpSorting,
                  fields: [
                    {
                      id: "mpName",
                      title: strings.NameTitle,
                      type: this.customCollectionFieldType.string,
                      required: true,
                      onGetErrorMessage: this.validateSortingProperty,
                      deferredValidationTime: 500
                    },
                    {
                      id: "mpOrder",
                      title: strings.SortOrderTitle,
                      type: this.customCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "ascending",
                          text: strings.Ascending
                        },
                        {
                          key: "descending",
                          text: strings.Descending
                        }
                      ],
                      defaultValue: "ascending",
                      required: true
                    }
                  ],
                  disabled: false
                }),
                PropertyPaneToggle('duplicates', {
                  label: strings.DuplicatesFieldLabel,
                  onText: strings.DuplicatesFieldLabelOn,
                  offText: strings.DuplicatesFieldLabelOff
                }),
                PropertyPaneToggle('privateGroups', {
                  label: strings.PrivateGroupsFieldLabel,
                  onText: strings.PrivateGroupsFieldLabelOn,
                  offText: strings.PrivateGroupsFieldLabelOff
                })
              ],
              isCollapsed: true
            },
            {
              groupName: strings.TemplateGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneToggle('debug', {
                  label: strings.DebugFieldLabel,
                  onText: strings.DebugFieldLabelOn,
                  offText: strings.DebugFieldLabelOff
                }),
                PropertyPaneTextField('external', {
                  label: strings.ExternalFieldLabel,
                  onGetErrorMessage: this._externalTemplateValidation.bind(this)
                }),
                PropertyPaneToggle('scriptloading', {
                  label: strings.ScriptloadingFieldLabel,
                  onText: strings.ScriptloadingFieldLabelOn,
                  offText: strings.ScriptloadingFieldLabelOff
                })
              ],
              isCollapsed: true
            },
            {
              groupName: strings.AudienceGroupName,
              groupFields: [
                PropertyPaneTextField('audienceColumnMapping', {
                  label: strings.AudienceColumnMappingLabel,
                  description: strings.AudienceColumnMappingDescription,
                  multiline: true
                }),
                PropertyPaneDropdown('audienceBooleanOperator', {
                  label: strings.AudienceBooleanOperatorLabel,
                  ariaLabel: strings.AudienceBooleanOperatorLabel,
                  options: [
                    { key: 'OR', text: 'OR' },
                    { key: 'AND', text: 'AND' }
                  ],
                  selectedKey: 'OR',
                }),
                PropertyPaneTextField('audienceColumnAllValue', {
                  label: strings.AudienceAllValueLabel,
                  description: strings.AudienceAllValueDescription
                })
              ],
              isCollapsed: true
            }
          ],
          displayGroupsAsAccordion: true
        }
      ]
    };
  }

  /**
  * Returns the sorting options in the right format
  */
  private getSortingOption() {
    const { mpSorting } = this.properties;
    if (mpSorting && mpSorting.length > 0) {
      return mpSorting.map(mp => `${mp.mpName}:${mp.mpOrder}`).join(',');
    }
    return null;
  }

  /**
  * Check if the provided managed property is sortable
  *
  * @param value
  * @param index
  * @param crntItem
  */
  private validateSortingProperty = async (value: any): Promise<string> => {
    if (value) {
      try {
        const searchApi = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='*'&sortlist='${value}:ascending'&RowLimit=1&selectproperties='Path'`;
        const data = await this.context.spHttpClient.get(searchApi, SPHttpClient.configurations.v1);
        return data.ok ? "" : Text.format(strings.InvalidSortingFieldDescription, value);
      } catch (e) {
        console.log(e);
        return "Something failed";
      }
    }
    return "";
  }

  /**
  * Validating the query property
  *
  * @param value
  */
  private _queryValidation(value: string): string {
    // Check if a URL is specified
    if (value.trim() === "") {
      return strings.QuertValidationEmpty;
    }

    return '';
  }

  /**
  * Validating the external template property
  *
  * @param value
  */
  private _externalTemplateValidation(value: string): string {
    // If debug template is set to off, user needs to specify a template URL
    if (!this.properties.debug) {
      // Check if a URL is specified
      if (value.trim() === "") {
        return strings.TemplateValidationEmpty;
      }
      // Check if a HTML file is referenced
      if (value.toLowerCase().indexOf('.html') === -1) {
        return strings.TemplateValidationHTML;
      }
    }

    return '';
  }

  /**
  * Prevent from changing the query on typing
  */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  /**
  * Retrieves user profile properties
  */
  private _getUserProfileProperties(): Promise<any> {
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/sp.userprofiles.peoplemanager/getmyproperties`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    }).catch(error => {
      return Promise.reject(JSON.stringify(error));
    });
  }
}
