import { ISearchResults, ICells, ICellValue, ISearchResponse } from './ISearchService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import SearchTokenHelper from "../helpers/SearchTokenHelper";

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
    public get(query: string, maxResults: number, sorting: string, fields: string[] = []) {
        return new Promise<ISearchResponse>((resolve, reject) => {
            let url: string = this._context.pageContext.web.absoluteUrl + "/_api/search/query?querytext=";
            // Check if a query is provided
            url += !this._isEmptyString(query) ? `'${this._tokenHelper.replaceTokens(query)}'` : "'*'";
            // Check if there are fields provided
            if (!this._isEmptyString(fields.join(','))) {
                url += `&selectproperties='${fields}'`;
            }
            // Add the rowlimit
            url += "&rowlimit=";
            url += !this._isNull(maxResults) ? maxResults : 3;
            // Add sorting
            url += !this._isEmptyString(sorting) ? `&sortlist='${sorting}'` : "";
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
                    searchUrl: url
                };
                resolve(searchResp);
            }).catch((error: string) => reject(error));
        });
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
