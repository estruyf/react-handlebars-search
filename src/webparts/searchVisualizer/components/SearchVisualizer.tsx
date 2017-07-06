import * as React from 'react';
import * as Handlebars from 'handlebars';

import styles from './SearchVisualizer.module.scss';
import { ISearchVisualizerProps, ISearchVisualizerState, IMetadata } from './ISearchVisualizerProps';
import { SPHttpClient } from "@microsoft/sp-http";
import SPHttpClientResponse from "@microsoft/sp-http/lib/spHttpClient/SPHttpClientResponse";
import executeScript from "../helpers/DangerousScriptLoader";
import TypeofHelper from "../helpers/TypeofHelper";
import SearchService from "../services/SearchService";
import { Spinner, SpinnerSize, MessageBar, MessageBarType, Dialog, DialogType } from 'office-ui-fabric-react';
import { ISearchResponse } from "../services/ISearchService";
import * as strings from 'searchVisualizerStrings';


export interface SPUser {
    username?: string;
    displayName?: string;
    email?: string;
}

export default class SearchVisualizer extends React.Component<ISearchVisualizerProps, ISearchVisualizerState> {
    private _searchService: SearchService;
    private _results: any[] = [];
    private _fields: string[] = [];
    private _templateMarkup: string = "";
    private _tmplDoc: Document;

    constructor(props: ISearchVisualizerProps, state: ISearchVisualizerState) {
        super(props);

        // Initialize the search service
        this._searchService = new SearchService(props.context);

        // Initialize the current component state
        this.state = {
            loading: true,
            template: "",
            error: "",
            showError: false,
            showScriptDialog: false
        };

        // Bind "this" to the load template function
        this._loadTemplate = this._loadTemplate.bind(this);
        this._processResults = this._processResults.bind(this);

        // Load all the handlebars helpers
        let helpers = require<any>('handlebars-helpers')({
            handlebars: Handlebars
        });

        // Load the typeof field handler for debugging
        Handlebars.registerHelper('typeof', TypeofHelper);


        // SharePoint helper to split SPUserField (?multiple) into a string
        // The template provide the property which will be returned
        Handlebars.registerHelper('splitSPUser', function (userFieldValue, propertyRequested) {

            if (userFieldValue == null)
                return null;

            const retValue:string[]=[];
            let userFieldValueArray = userFieldValue.split(';').forEach(user => {
                let userValues = user.split(' | ');
                let spuser: SPUser = {
                    displayName: userValues[1],
                    email: userValues[0]
                }
                retValue.push(spuser[propertyRequested]);
            });

            return retValue.join(', ');
        });

        // SHarePoint helper to split the displaynames of for example the Author field (user1;user2...)
        Handlebars.registerHelper('splitDisplayNames', function (displayNames) {

            if (displayNames == null && displayNames.indexOf(';') == -1)
                return null;

            return displayNames.split(';').join(", ");;
        });
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
            prevProps.sorting !== this.props.sorting) {
            this._resetLoadingState();
            // Only refresh the search results
            this._processResults();
        }
    }

    private _resetLoadingState() {
        this.setState({
            loading: true,
            error: "",
            showError: false
        });
    }

    /**
     * Processing the search web part tasks
     */
    private _processSearchTasks(): void {
        this._loadTemplate()
            .then((tmpl: string) => {
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
            })
            .catch((error: string) => {
                this.setState({
                    error: error
                });
            });
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
    private _loadTemplate(): Promise<string> {
        // Check if internal template must be used or when debugging is turned on
        if (!this.props.external || this.props.debug) {
            return Promise.resolve(require('./debug.template.html'));
        }

        return new Promise((resolve, reject) => {
            this.props.context.spHttpClient.get(this.props.external, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    resolve(response.text());
                } else {
                    reject(`Template: ${response.statusText}`);
                }
            });
        });
    }

    /**
     * Processing the search result retrieval process
     */
    private _processResults() {
        //  Get the search results and then bind it to the template
        this._searchService.get(this.props.query, this.props.maxResults, this.props.sorting, this._fields).then((searchResp: ISearchResponse) => {
            // Create the template values object
            const tmplValues: any = {
                wpTitle: this.props.title,
                pageCtx: this.props.context.pageContext,
                items: searchResp.results,
                calledUrl: searchResp.searchUrl
            };

            // Reload the new template
            let template: HandlebarsTemplateDelegate = Handlebars.compile(this._templateMarkup);
            let templateResult = template(tmplValues);

            // Update the current component state
            this.setState({
                loading: false,
                template: templateResult
            });

            // Check if the wp needs to execute the scripts in the HTML
            if (this.props.scriptloading) {
                executeScript(this._tmplDoc.getElementById('template'));
            }
        }).catch((error: any) => {
            this.setState({
                error: error.toString()
            });
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
                <MessageBar className={styles.error} messageBarType={ MessageBarType.error }>
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
            <div className={styles.searchVisualizer}>
                {view}

                <Dialog isOpen={this.state.showScriptDialog} type={DialogType.normal} onDismiss={this._toggleDialog.bind(this)} title={strings.ScriptsDialogHeader} subText={strings.ScriptsDialogSubText}></Dialog>
            </div>
        );
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
}
