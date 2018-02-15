import * as Handlebars from 'handlebars';

/**
 * Various HB helpers
 */
export default class HBVariousHelpers {
  public static init() {
    // Register the helpers
    Handlebars.registerHelper('typeof', this._typeof);
  }

  /**
   * Return the type of from the object which is passed
   * @param context
   * @param options
   */
  private static _typeof(context, options) {
    return typeof context === "object" ? Object.prototype.toString.call(context) : typeof context;
  }
}
