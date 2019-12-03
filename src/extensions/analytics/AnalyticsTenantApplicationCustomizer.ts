import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { Log } from '@microsoft/sp-core-library';

const LOG_SOURCE: string = 'GoogleAnalyticsTenantApplicationCustomizer';

export interface IAnalyticsTenantApplicationCustomizerProperties {
  googleTrackingId: string;
}

export default class AnalyticsTenantApplicationCustomizer
  extends BaseApplicationCustomizer<IAnalyticsTenantApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    try {
      let googleTrackingId: string = this.properties.googleTrackingId;

      //If we have a google tracking id, setup ga on page
      if (googleTrackingId) {
        Log.verbose(LOG_SOURCE, `Google Analytics Tracking ID: ${googleTrackingId}`);
        this.setupGoogleAnalytics(googleTrackingId);
      }
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
  }

  private setupGoogleAnalytics(googleTrackingId: string): void {
    //Add Google Analytics Script Tag to Page
    var gtagScript = document.createElement("script");
    gtagScript.type = "text/javascript";
    gtagScript.src = `https://www.googletagmanager.com/gtag/js?id=${googleTrackingId}`;
    gtagScript.async = true;
    document.head.appendChild(gtagScript);

    //Invoke Google Analytics Page Tracker
    window["dataLayer"] = window["dataLayer"] || [];
    /* tslint:disable:no-function-expression */
    window["gtag"] = window["gtag"] || function gtag() {
      window["dataLayer"].push(arguments);
    };
    /* tslint:enable:no-function-expression */
    window["gtag"]('js', new Date());
    window["gtag"]('config', googleTrackingId);
  }

}
