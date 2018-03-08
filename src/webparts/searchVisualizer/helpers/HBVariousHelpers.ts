import * as Handlebars from 'handlebars';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { toArray } from '@microsoft/sp-lodash-subset';

/**
 * Various HB helpers
 */
export default class HBVariousHelpers {
    constructor(private _context: WebPartContext) {
        Handlebars.registerHelper('typeof', this._typeof);
        Handlebars.registerHelper('log', this._log);
    }

    /**
       * Initialize the class
       * @param _context
       */
    public static init(_context: WebPartContext) {
        const instance = new HBVariousHelpers(_context);
    }

    /**
     * Return the type of from the object which is passed
     * @param context
     * @param options
     */
    private _typeof(context, options) {
        return typeof context === "object" ? Object.prototype.toString.call(context) : typeof context;
    }

    private _log(level) {
        console.log.apply(console, [].concat(["Handlebars: "], toArray(arguments)));
    }

}
