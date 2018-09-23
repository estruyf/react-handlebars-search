import { SPHttpClient } from "@microsoft/sp-http";

export class UserProfileService {
  public static readonly USERPROFILE_KEY = 'SearchVisualizerWebPart:UserProfileData';

  public static async getProperties(spHttpClient: SPHttpClient, webUrl: string) {
    try {
      const userProfileData = window.sessionStorage ? sessionStorage.getItem(UserProfileService.USERPROFILE_KEY) : null;
      if (userProfileData) {
        return userProfileData;
      }

      const apiUrl = `${webUrl}/_api/sp.userprofiles.peoplemanager/getmyproperties`;
      const resp = await spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      if (resp.ok) {
        const userData = await resp.json();
        if (userData.UserProfileProperties) {
          sessionStorage.setItem(UserProfileService.USERPROFILE_KEY, JSON.stringify(userData.UserProfileProperties));
        }
        return userData.UserProfileProperties || null;
      } else {
        return null;
      }
    } catch (e) {
      console.error(`Sorry, something failed while fetching your user profile properties.`);
      console.error(e);
      return null;
    }
  }
}
