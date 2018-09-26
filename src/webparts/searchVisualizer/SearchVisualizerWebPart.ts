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
import { SPHttpClient } from '@microsoft/sp-http';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { SearchFilter, IAdvancedFilter } from './models/IAdvancedFilter';
import { UserProfileService } from './services/UserProfileService';

export default class SearchVisualizerWebPart extends BaseClientSideWebPart<ISearchVisualizerWebPartProps> {

  /**
  * Renders the search visualizer component
  */
  public render(): void {
    // Load the core UI Fabric styles
    SPComponentLoader.loadCss('https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.0/css/fabric-9.6.0.scoped.min.css');

    const element: React.ReactElement<ISearchVisualizerProps> = React.createElement(
      SearchVisualizer, {
        ...this.properties,
        sorting: this.getSortingOption(),
        audienceTargetingBooleanOperator: this.properties.audienceTargetingBooleanOperator ? this.properties.audienceTargetingBooleanOperator : 'OR',
        context: this.context,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    const domElement: HTMLElement = this.domElement;
    if (this.properties.audienceTargeting && this.properties.audienceTargetingAll && window.sessionStorage) {
      UserProfileService.getProperties(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl).then(() => {
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
  * Webpart Property Pane Configuration
  */
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
                  onGetErrorMessage: this.queryValidation,
                  deferredValidationTime: 500,
                  disabled: this.properties.advancedSearch && this.properties.advancedSearch.length > 0
                }),
                PropertyPaneSlider('maxResults', {
                  label: strings.FieldsMaxResults,
                  min: 1,
                  max: 50
                }),
                PropertyPaneTextField('sorting', {
                  label: strings.SortingFieldLabel
                })
              ],
              isCollapsed: false
            }
          ],
          displayGroupsAsAccordion: true
        }
      ]
    };
  }

  /**
  * Validating the query property
  *
  * @param value
  */
  private queryValidation(value: string): string {
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
  private externalTemplateValidation(value: string): string {
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
  * Property pane field changes
  *
  * @param propertyPath
  * @param oldValue
  * @param newValue
  */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    if (propertyPath === "advancedSearch") {
      // Remove the old value
      if (oldValue && oldValue.length > 0) {
        const oldQuery = oldValue.map(q => this.createFilterQuery(q as IAdvancedFilter)).join(" ");
        this.properties.query = this.properties.query.replace(oldQuery, "").trim();
      }

      // Add the new values
      if (newValue && newValue.length > 0) {
        const advancedQuery = newValue.map(q => this.createFilterQuery(q as IAdvancedFilter)).join(" ");
        this.properties.query = `${this.properties.query} ${advancedQuery}`;
      }
    }
  }

  /**
  * Create the keyword filter query for the provided row
  *
  * @param item
  */
  private createFilterQuery(item: IAdvancedFilter) {
    switch (item.filter) {
      case SearchFilter.containsStartsWith:
      return `${item.name}:${item.value}* ${item.operator || ""}`.trim();
      case SearchFilter.startsWith:
      return `${item.name}=${item.value}* ${item.operator || ""}`.trim();
      case SearchFilter.notContains:
      return `-${item.name}:${item.value} ${item.operator || ""}`.trim();
      case SearchFilter.notStartsWith:
      return `-${item.name}=${item.value}* ${item.operator || ""}`.trim();
      default:
      return `${item.name}${item.filter}${item.value} ${item.operator || ""}`.trim();
    }
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
  private validateSortingProperty = async (value: any, index: number, crntItem: any): Promise<string> => {
    if (value) {
      try {
        const searchApi = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='*'&sortlist='${value}:ascending'&RowLimit=1&selectproperties='Path'`;
        const data = await this.context.spHttpClient.get(searchApi, SPHttpClient.configurations.v1);
        return data.ok ? "" : `The managed property "${value}" is not sortable.`;
      } catch (e) {
        console.log(e);
        return "Something failed";
      }
    }
    return "";
  }
}
