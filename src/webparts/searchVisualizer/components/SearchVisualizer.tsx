import * as React from 'react';
import * as Handlebars from 'handlebars';

import styles from './SearchVisualizer.module.scss';
import { ISearchVisualizerProps, ISearchVisualizerState, IMetadata, ITemplateResource } from './ISearchVisualizerProps';
import { SPHttpClient } from "@microsoft/sp-http";
import executeScript from "../helpers/DangerousScriptLoader";
import SearchService from "../services/SearchService";
import { Spinner, SpinnerSize, MessageBar, MessageBarType, Dialog, DialogType } from 'office-ui-fabric-react';
import { ISearchResponse } from "../services/ISearchService";
import * as strings from 'searchVisualizerStrings';
import * as uuidv4 from 'uuid/v4';
import CustomHelpers from '../helpers/CustomHelpers';


export default class SearchVisualizer extends React.Component<ISearchVisualizerProps, ISearchVisualizerState> {
    private _searchService: SearchService;
    private _results: any[] = [];
    private _fields: string[] = [];
    private _templateMarkup: string = "";
    private _tmplDoc: Document;
    private _totalResults: number = 0;
    private _pageNr: number = 0;
    private _compId: string = "";
    private _resources: ITemplateResource[] = [];

    constructor(props: ISearchVisualizerProps, state: ISearchVisualizerState) {
        super(props);

        // Initialize the search service
        this._searchService = new SearchService(props.context);

        // Specify a unique ID for the component
        this._compId = 'search-' + uuidv4();

        // Initialize the current component state
        this.state = {
            loading: true,
            template: "",
            error: "",
            showError: false,
            showScriptDialog: false
        };

        // Load all the handlebars helpers
        let helpers = require<any>('handlebars-helpers')({
            handlebars: Handlebars
        });

        // Register all custom helpers
        CustomHelpers.init(this.props.context);
    }

    /**
     * Called after initial render
     */
    public componentDidMount(): void {
        this._processSearchTasks();
    }

    /**
     * Called after a properties or state update
     * @param prevProps
     * @param prevState
     */
    public componentDidUpdate(prevProps: ISearchVisualizerProps, prevState: ISearchVisualizerState): void {
        // Check if the template needs to be updated
        if (prevProps.title !== this.props.title ||
            prevProps.debug !== this.props.debug ||
            prevProps.external !== this.props.external ||
            prevProps.scriptloading !== this.props.scriptloading) {
            this._resetLoadingState();
            // Refresh template and search results
            this._processSearchTasks();
        } else if (prevProps.query !== this.props.query ||
            prevProps.maxResults !== this.props.maxResults ||
            prevProps.sorting !== this.props.sorting ||
            prevProps.duplicates !== this.props.duplicates ||
            prevProps.privateGroups !== this.props.privateGroups) {
            this._resetLoadingState();
            // Only refresh the search results
            this._processResults();
        }
    }

    private _resetLoadingState() {
        // Reset page number
        this._pageNr = 0;
        // Reset state
        this.setState({
            loading: true,
            error: "",
            showError: false
        });
    }

    /**
     * Processing the search web part tasks
     */
    private async _processSearchTasks(): Promise<void> {
      const tmpl: string | null = await this._loadTemplate();
      if (tmpl) {
        // Parse the template
        const parser = new DOMParser();
        this._tmplDoc = parser.parseFromString(tmpl, 'text/html');

        // Get the field metadata
        let metadata: IMetadata = JSON.parse(this._tmplDoc.getElementById('metadata').innerHTML);
        if (metadata !== null) {
            if (typeof metadata.fields !== "undefined") {
                this._fields = metadata.fields;
            } else {
                this._setDefaultMetadata();
            }

            // Check if the metadata contains resources
            if (typeof metadata.resources !== "undefined") {
                this._resources = metadata.resources;
            }
        } else {
            this._setDefaultMetadata();
        }

        // Get the template metadata
        this._templateMarkup = this._tmplDoc.getElementById('template').innerHTML;
        // When property pane is open, check if there are script tags in the provided template
        if (this.props.context.propertyPane.isPropertyPaneOpen() && !this.props.scriptloading) {
            if (this._templateMarkup.indexOf('<script') !== -1) {
                // Alert the user
                this.setState({
                    showScriptDialog: true
                });
            }
        }

        // Retrieve the next set of results
        this._processResults();
      }
    }

    /**
     * Set default metadata if not provided by the template
     */
    private _setDefaultMetadata() {
        this._fields = [];
    }

    /**
     * Loads the template for the web part
     */
    private _loadTemplate = async (): Promise<string> => {
      try {
        // Check if internal template must be used or when debugging is turned on
        if (!this.props.external || this.props.debug) {
          return require('./debug.template.html');
        }

        const response = await this.props.context.spHttpClient.get(this.props.external, SPHttpClient.configurations.v1);
        if (response.ok) {
          return response.text();
        } else {
          this.setState({
            error: `Template: ${response.statusText}`
          });
          return null;
        }
      } catch (e) {
        this.setState({
          error: `Failed retrieving the template`
        });
        return null;
      }
    }

    /**
     * Processing the search result retrieval process
     */
    private _processResults = async () => {
      const startRow = this._pageNr * this.props.maxResults;
      //  Get the search results and then bind it to the template
      try {
        const searchResp: ISearchResponse = await this._searchService.get(this.props.query, this.props.audienceTargeting, this.props.audienceTargetingAll, this.props.audienceTargetingBooleanOperator, this.props.maxResults, this.props.sorting, this.props.duplicates, this.props.privateGroups, startRow, this._fields);

        // Check which resources have to be loaded
        const locale = this.props.context.pageContext.cultureInfo.currentUICultureName;

        // Get tenant URL
        let tenantUrl = window.location.protocol + "//" + window.location.host;

        // Create a new resources object
        const resources = {};
        this._resources.forEach(resource => {
          if (resource.key && resource.values) {
            let value = "";
            // Check if it contains a default value
            if (resource.values["default"]) {
              value = resource.values["default"];
            }
            // Check if a resource value exists for the current language
            if (resource.values[locale.toLowerCase()]) {
              value = resource.values[locale.toLowerCase()];
            }
            // Set the resource value
            resources[resource.key] = value;
          }
        });

        // Create the template values object
        const tmplValues: any = {
          wpTitle: this.props.title,
          pageCtx: this.props.context.pageContext,
          items: searchResp.results,
          totalResults: searchResp.totalResults,
          totalResultsIncDuplicates: searchResp.totalResultsIncludingDuplicates,
          calledUrl: searchResp.searchUrl,
          resources: resources,
          locale: locale,
          tenantUrl: tenantUrl
        };

        // Reload the new template
        let template: HandlebarsTemplateDelegate = Handlebars.compile(this._templateMarkup);
        let templateResult = template(tmplValues);

        // Internally store the total results number
        this._totalResults = searchResp.totalResults;

        // Update the current component state
        this.setState({
          loading: false,
          template: templateResult
        });

        // Check if the wp needs to execute the scripts in the HTML
        if (this.props.scriptloading) {
          executeScript(this._tmplDoc.getElementById('template'));
        }

        // Bind the paging events
        this._bindPaging();
      } catch (e) {
        this.setState({
            error: e.toString()
        });
      }
    }


    /**
     * Bind the next and previous paging events to the paging elements defined in the template
     */
    private _bindPaging() {
        const prevPageElm = document.querySelector(`#${this._compId} #prevPage`);
        const nextPageElm = document.querySelector(`#${this._compId} #nextPage`);

        if (prevPageElm) {
            // Check if the element needs to be disabled
            if (this._pageNr <= 0) {
                prevPageElm.classList.add('disabled');
            } else {
                prevPageElm.classList.remove('disabled');
                prevPageElm.addEventListener("click", () => {
                    this._prevPage();
                });
            }
        }

        if (nextPageElm) {
            // Check if the element needs to be disabled
            if (this._totalResults > (this.props.maxResults * (this._pageNr + 1))) {
                nextPageElm.classList.remove('disabled');
                nextPageElm.addEventListener("click", () => {
                    this._nextPage();
                });
            } else {
                nextPageElm.classList.add('disabled');
            }
        }
    }

    /**
     * Get the results of the previous page
     */
    private _prevPage = () => {
        this._pageNr--;
        this._processResults();
    }

    /**
     * Get the results of the next page
     */
    private _nextPage = () => {
        this._pageNr++;
        this._processResults();
    }


    /**
     * Toggle the show error message
     */
    private _toggleError() {
        this.setState({
            showError: !this.state.showError
        });
    }

    /**
     * Toggle the script dialog visibility
     */
    private _toggleDialog() {
        this.setState({
            showScriptDialog: !this.state.showScriptDialog
        });
    }

    /**
     * Render the contents of the web part
     */
    public render(): React.ReactElement<ISearchVisualizerProps> {
        let view = <Spinner size={SpinnerSize.large} label='Loading results' />;

        if (!this.state.loading && this.state.template) {
            view = <div dangerouslySetInnerHTML={{ __html: this.state.template }}></div>;
        }

        if (this.state.error !== "") {
            return (
                <MessageBar className={styles.error} messageBarType={MessageBarType.error}>
                    <span>Sorry, something went wrong</span>
                    {
                        (() => {
                            if (this.state.showError) {
                                return (
                                    <div>
                                        <p>
                                            <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m"><i className={`ms-Icon ms-Icon--ChevronUp ${styles.icon}`} aria-hidden="true"></i> Hide the error message</a>
                                        </p>
                                        <p className="ms-font-m">{this.state.error}</p>
                                    </div>
                                );
                            } else {
                                return (
                                    <p>
                                        <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m"><i className={`ms-Icon ms-Icon--ChevronDown ${styles.icon}`} aria-hidden="true"></i> Show the error message</a>
                                    </p>
                                );
                            }
                        })()
                    }
                </MessageBar>
            );
        }

        return (
            <div id={this._compId} className={`${styles.searchVisualizer} ms-Fabric--v6-0-0`}>
                {view}

                <Dialog isOpen={this.state.showScriptDialog} type={DialogType.normal} onDismiss={this._toggleDialog.bind(this)} title={strings.ScriptsDialogHeader} subText={strings.ScriptsDialogSubText}></Dialog>
            </div>
        );
    }
}
