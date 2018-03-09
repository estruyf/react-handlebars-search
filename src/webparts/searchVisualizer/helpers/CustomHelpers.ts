import { WebPartContext } from '@microsoft/sp-webpart-base';
import HBSharePointHelpers from "./HBSharePointHelpers";
import HBVariousHelpers from './HBVariousHelpers';

export default class CustomHelpers {
  public static init(_context: WebPartContext): void {
    // Register various helpers
    HBVariousHelpers.init();
    // Register the SharePoint helpers
    HBSharePointHelpers.init(_context);
  }
}
