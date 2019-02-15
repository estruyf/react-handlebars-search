import { IUserProfileProperty } from './IUserProfileProperty';
import { USERPROFILE_KEY } from './../SearchVisualizerWebPart';
import { ISearchResults, ICells, ICellValue, ISearchResponse } from './ISearchService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import SearchTokenHelper from "../helpers/SearchTokenHelper";
import { IAudienceProperty } from './IAudienceProperty';

export default class SearchService {
  private _tokenHelper: SearchTokenHelper;
  private _results: any[];

  constructor(private _context: IWebPartContext) {
    this._tokenHelper = new SearchTokenHelper(_context);
  }

  /**
  * Retrieve search results by the specified query
  *
  * @param query
  * @param maxResults
  * @param sorting
  * @param fields
  */
  public get(query: string, audienceTargeting: string, audienceTargetingAll: string, audienceTargetingBooleanOperator: string, maxResults: number, sorting: string, duplicates: boolean, privateGroups: boolean, startRow: number, fields: string[] = []): Promise<ISearchResponse> {
    return new Promise<ISearchResponse>((resolve, reject) => {
      let totalResults: number = null;
      let totalRowsIncludingDuplicates: number = null;

      let url: string = this._context.pageContext.web.absoluteUrl + "/_api/search/query?querytext=";

      // check for audience targeting
      let audienceQuery: string = '';
      if (!this._isEmptyString(audienceTargeting) && !this._isEmptyString(audienceTargetingAll)) {
        audienceQuery = this.BuildAudienceQuery(audienceTargeting, audienceTargetingAll, audienceTargetingBooleanOperator);
      }

      // Check if a query is provided
      url += !this._isEmptyString(query) ? `'${encodeURIComponent(this._tokenHelper.replaceTokens(query))} ${audienceQuery}'` : "'*'";

      // Check if there are fields provided
      if (!this._isEmptyString(fields.join(','))) {
        url += `&selectproperties='${fields}'`;
      }
      // Add the rowlimit
      url += "&rowlimit=";
      url += !this._isNull(maxResults) ? maxResults : 3;
      // Check the startrow
      url += startRow <= 0 ? "" : `&startrow=${startRow}`;
      // Add sorting
      url += !this._isEmptyString(sorting) ? `&sortlist='${sorting}'` : "";
      // Check if result duplicates needs to get trimmed
      url += !duplicates ? "&trimduplicates=false" : "";
      // Check if the user wants to search for private group data
      url += privateGroups ? "&Properties='EnableDynamicGroups:true'" : "";
      // Add the client type
      url += "&clienttype='ContentSearchRegular'";
      // Do an Ajax call to receive the search results
      this._getSearchData(url).then((res: ISearchResults) => {
        // Check if there was an error
        if (typeof res["odata.error"] !== "undefined") {
          if (typeof res["odata.error"]["message"] !== "undefined") {
            reject(res["odata.error"]["message"].value);
            return;
          }
        }

        let resultsRetrieved = false;
        // Check if results were retrieved
        if (!this._isNull(res)) {
          if (typeof res.PrimaryQueryResult !== 'undefined') {
            if (typeof res.PrimaryQueryResult.RelevantResults !== 'undefined') {
              // Retrieve the total rows number
              if (res.PrimaryQueryResult.RelevantResults.TotalRows) {
                totalResults = res.PrimaryQueryResult.RelevantResults.TotalRows;
              }
              // Retrieve the total rows including the duplicates
              if (res.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates) {
                totalRowsIncludingDuplicates = res.PrimaryQueryResult.RelevantResults.TotalRowsIncludingDuplicates;
              }
              // Retrieve all the table rows
              if (typeof res.PrimaryQueryResult.RelevantResults.Table !== 'undefined') {
                if (typeof res.PrimaryQueryResult.RelevantResults.Table.Rows !== 'undefined') {
                  resultsRetrieved = true;
                  this._setSearchResults(res.PrimaryQueryResult.RelevantResults.Table.Rows, fields.join(','));
                }
              }
            }
          }
        }

        // Reset the store its search result set on error
        if (!resultsRetrieved) {
          this._setSearchResults([], null);
        }

        // Return the retrieved result set
        const searchResp: ISearchResponse = {
          results: this._results,
          totalResults: totalResults,
          totalResultsIncludingDuplicates: totalRowsIncludingDuplicates,
          searchUrl: url
        };
        resolve(searchResp);
      }).catch((error: string) => reject(error));
    });
  }

  /**
  * Adds audience targetting support to the query
  *
  * @param audienceTargeting
  * @param audienceTargetingAll
  * @param audienceTargetingBooleanOperator
  */
  private BuildAudienceQuery(audienceTargeting: string, audienceTargetingAll: string, audienceTargetingBooleanOperator: string): string {
    // Check session storage for user profile data
    if (window.sessionStorage) {
      const userProfileData = sessionStorage.getItem(USERPROFILE_KEY);
      if (userProfileData) {
        let properties: IUserProfileProperty[] = JSON.parse(userProfileData);

        let columnMapping: IAudienceProperty = JSON.parse(audienceTargetingAll);
        let managedPropertyName: string = Object.keys(columnMapping)[0];
        let managedPropertyValue: string = columnMapping[managedPropertyName];

        let baseAudienceQuery: string = `${managedPropertyName}="${managedPropertyValue}"`;

        let targets: string[] = audienceTargeting.split('\n');

        let audienceQuery: string = '';
        for (let i: number = 0, max = targets.length; i < max; i++) {
          columnMapping = JSON.parse(targets[i]);

          managedPropertyName = Object.keys(columnMapping)[0];
          let userProfilePropertyName: string = columnMapping[managedPropertyName];

          let property: IUserProfileProperty = this.FilterUserProfileProperties(properties, userProfilePropertyName);
          if (property && property.Value) {
            audienceQuery = `${audienceQuery}${managedPropertyName}="${property.Value}"`;

            if (i + 1 < max) {
              audienceQuery = `${audienceQuery} ${audienceTargetingBooleanOperator} `;
            } else {
              audienceQuery = `${audienceQuery}`;
            }
          }
        }
        return audienceQuery ? `(${baseAudienceQuery} OR (${audienceQuery}))` : `(${baseAudienceQuery})`;
      }
    }
    return "";
  }


  /**
  * Retrieves the User Profile property based on the key value
  *
  * @param properties
  * @param propertyName
  */
  private FilterUserProfileProperties(properties: IUserProfileProperty[], propertyName: string): IUserProfileProperty {
    for (var i = 0, len = properties.length; i < len; i++) {
      if (properties[i].Key === propertyName) {
        return properties[i];
      }
    }
    return null;
  }


  /**
  * Retrieve the results from the search API
  *
  * @param context
  * @param url
  */
  private _getSearchData(url: string): Promise<ISearchResults> {
    return this._context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        'odata-version': '3.0'
      }
    }).then((res: SPHttpClientResponse) => {
      return res.json();
    }).catch(error => {
      return Promise.reject(JSON.stringify(error));
    });
  }

  /**
  * Set the current set of search results
  *
  * @param crntResults
  * @param fields
  */
  private _setSearchResults(crntResults: ICells[], fields: string): void {
    if (crntResults.length > 0) {
      const flds: string[] = fields.toLowerCase().split(',');
      const temp: any[] = [];
      crntResults.forEach((result) => {
        // Create a temp value
        var val: Object = {};
        result.Cells.forEach((cell: ICellValue) => {
          if (flds.indexOf(cell.Key.toLowerCase()) !== -1) {
            // Add key and value to temp value
            val[cell.Key] = cell.Value;
          } else if (flds.length === 1 && flds[0] === "") { // If empty fields variable, return all fields
            // Add key and value to temp value
            val[cell.Key] = cell.Value;
          }
        });
        // Push this to the temp array
        temp.push(val);
      });
      this._results = temp;
    } else {
      this._results = [];
    }
  }

  /**
  * Check if the value is null, undefined or empty
  *
  * @param value
  */
  private _isEmptyString(value: string): boolean {
    return value === null || typeof value === "undefined" || !value.length;
  }

  /**
  * Check if the value is null or undefined
  *
  * @param value
  */
  private _isNull(value: any): boolean {
    return value === null || typeof value === "undefined";
  }
}
