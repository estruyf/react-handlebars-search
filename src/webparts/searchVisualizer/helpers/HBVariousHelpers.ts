import * as Handlebars from 'handlebars';

/**
 * Various HB helpers
 */
export default class HBVariousHelpers {
  public static init() {
    // Register the helpers
    Handlebars.registerHelper('typeof', this._typeof);
    Handlebars.registerHelper('fileExtIcon', this.fileExtIcon);
  }

  /**
   * Return the type of from the object which is passed
   * @param context
   * @param options
   */
  private static _typeof(context, options) {
    return typeof context === "object" ? Object.prototype.toString.call(context) : typeof context;
  }

  private static fileExtIcon(fileExt: string) {
    switch (fileExt.toLowerCase()) {
      case "aspx":
        return "https://spoprod-a.akamaihd.net/files/odsp-next-prod_2018-09-07-sts_20180920.002/odsp-media/images/itemtypes/20_2x/html.png";
      default:
        return `https://spoprod-a.akamaihd.net/files/odsp-next-prod_2018-09-07-sts_20180920.002/odsp-media/images/itemtypes/20_2x/${fileExt}.png`;
    }
  }
}
