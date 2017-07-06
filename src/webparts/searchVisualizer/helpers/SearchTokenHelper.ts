import { IWebPartContext } from '@microsoft/sp-webpart-base';

import * as moment from 'moment';

export default class SearchTokenHelper {
    private regexVal: RegExp = /\{[^\{]*?\}/gi;

    constructor(private _context: IWebPartContext) {}

    public replaceTokens(restUrl: string): string {
        const tokens = restUrl.match(this.regexVal);

        if (tokens !== null && tokens.length > 0) {
            tokens.forEach((token) => {
                // Check which token has been retrieved
                if (token.toLowerCase().indexOf('today') !== -1) {
                    const dateValue = this.getDateValue(token);
                    restUrl = restUrl.replace(token, dateValue);
                }
                else if (token.toLowerCase().indexOf('user') !== -1) {
                    const userValue = this.getUserValue(token);
                    restUrl = restUrl.replace(token, userValue);
                }
                else {
                    switch (token.toLowerCase()) {
                        case "{site}":
                            restUrl = restUrl.replace(/{site}/ig, this._context.pageContext.web.absoluteUrl);
                            break;
                        case "{sitecollection}":
                            restUrl = restUrl.replace(/{sitecollection}/ig, this._context.pageContext.site.absoluteUrl);
                            break;
                        case "{currentdisplaylanguage}":
                            restUrl = restUrl.replace(/{currentdisplaylanguage}/ig, this._context.pageContext.cultureInfo.currentCultureName);
                            break;
                    }
                }
            });
        }

		return restUrl;
    }

    private getDateValue(token: string): string {
        let dateValue = moment();
        // Check if we need to add days
        if (token.toLowerCase().indexOf("{today+") !== -1) {
            const daysVal = this.getDaysVal(token);
            dateValue = dateValue.add(daysVal, 'day');
        }
        // Check if we need to subtract days
        if (token.toLowerCase().indexOf("{today-") !== -1) {
            const daysVal = this.getDaysVal(token);
            dateValue = dateValue.subtract(daysVal, 'day');
        }
        return dateValue.format('YYYY-MM-DD');
    }

    private getDaysVal(token: string): number {
        const tmpDays: string = token.substring(7, token.length - 1);
        return parseInt(tmpDays) || 0;
    }

    private getUserValue(token: string): string {
        let userValue = `"${this._context.pageContext.user.displayName}"`;

        if (token.toLowerCase().indexOf("{user.") !== -1) {
            const propVal = token.toLowerCase().substring(6, token.length - 1);
            switch (propVal) {
                case "name":
                    userValue = `"${this._context.pageContext.user.displayName}"`;
                    break;
                case "email":
                    userValue = this._context.pageContext.user.email;
                    break;
            }
        }

        return userValue;
    }
}
