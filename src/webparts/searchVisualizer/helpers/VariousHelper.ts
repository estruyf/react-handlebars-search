import { IWebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Various helpers
 */
export default class VariousHelpers {
    constructor(private _context: IWebPartContext) {}

    /**
     * Return the URL Parameter
     * @param paramToRetrieve
     */
    public getQueryStringParameter(paramToRetrieve) {
        if (document.URL.split("?").length > 1) {
            let params = document.URL.split("?")[1].split("&");
            for (let i = 0; i < params.length; i = i + 1) {
                let singleParam = params[i].split("=");
                if (singleParam[0] == paramToRetrieve)
                    return decodeURIComponent(singleParam[1]);
            }
        }
    }
}
