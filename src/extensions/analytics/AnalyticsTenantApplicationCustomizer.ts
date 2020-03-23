import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { Log } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import sha256 from 'crypto-js/sha256';

const LOG_SOURCE: string = 'GoogleAnalyticsTenantApplicationCustomizer';

type PropertyMapping = {
  spoUserPropertyName: string;
  gtmPropertyName?: string;
  encryptSHA256?: boolean;
};

export interface IAnalyticsTenantApplicationCustomizerProperties {
  googleTrackingId: string;
  useGoogleTagManager?: boolean;
  propertyMappings?: PropertyMapping[];
  disableCache?: boolean;
  cacheLifetimeMins?: number;
}

const DEFAULT_PROPERTIES: IAnalyticsTenantApplicationCustomizerProperties = {
  googleTrackingId: "",
  useGoogleTagManager: false,
  propertyMappings: undefined,
  disableCache: false,
  cacheLifetimeMins: 60 * 12 /* 12 hours */
};

export default class AnalyticsTenantApplicationCustomizer
  extends BaseApplicationCustomizer<IAnalyticsTenantApplicationCustomizerProperties> {

  private getProperties(): IAnalyticsTenantApplicationCustomizerProperties {
    return { ...DEFAULT_PROPERTIES, ...this.properties };
  }

  @override
  public async onInit(): Promise<void> {
    try {
      let { googleTrackingId, useGoogleTagManager, propertyMappings } = this.getProperties();

      //If we have a google tracking id
      if (googleTrackingId) {
        Log.verbose(LOG_SOURCE, `Google Analytics Tracking ID: ${googleTrackingId}`);

        //Check if we should use Google Tag Manager
        if (useGoogleTagManager) {

          //If we have Property Mappings, fetch SharePoint User Profile and map to GTM data layer object
          if (propertyMappings) {
            const currentUserProfile = await this.getCurrentUserProperties();
            const dataLayerProperties = this.getDataLayerMappedProperties(propertyMappings, currentUserProfile);
            this.setupGoogleTagManager(googleTrackingId, dataLayerProperties);
          }
          //Else setup GTM without a data layer
          else {
            this.setupGoogleTagManager(googleTrackingId);
          }

        }
        //Else setup Google Analytics
        else {
          this.setupGoogleAnalytics(googleTrackingId);
        }
      }
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
  }

  private setupGoogleAnalytics(googleTrackingId: string): void {
    // Desconstructed Google Analytics gtag.js Embed Code
    // Docs: https://developers.google.com/analytics/devguides/collection/gtagjs

    var gaScript = document.createElement("script");
    gaScript.type = "text/javascript";
    gaScript.src = `https://www.googletagmanager.com/gtag/js?id=${googleTrackingId}`;
    gaScript.async = true;
    document.head.appendChild(gaScript);

    window["dataLayer"] = window["dataLayer"] || [];
    /* tslint:disable:no-function-expression */
    window["gtag"] = window["gtag"] || function gtag() {
      window["dataLayer"].push(arguments);
    };
    /* tslint:enable:no-function-expression */
    window["gtag"]('js', new Date());
    window["gtag"]('config', googleTrackingId);
  }

  private setupGoogleTagManager(googleTrackingId: string, dataLayerProperies?: object): void {
    // Deconstructed Google Tag Manager gtm.js Embed Code
    // Docs: https://developers.google.com/tag-manager/quickstart

    let customDataLayer = [];
    if (dataLayerProperies) {
      customDataLayer.push(dataLayerProperies);
    }
    customDataLayer.push({ 'gtm.start' : new Date().getTime(), event: 'gtm.js' });

    window["dataLayer"] = customDataLayer;

    var gtagScript = document.createElement('script');
    gtagScript.async = true;
    gtagScript.src = `https://www.googletagmanager.com/gtm.js?id=${googleTrackingId}`;
    const target = document.getElementsByTagName('script')[0];
    target.parentNode.insertBefore(gtagScript, target);
  }

  private getDataLayerMappedProperties(mappings: PropertyMapping[], userProfile: any): object {
    const dataLayerProperties = mappings.reduce((obj, propertyMapping) => {
      if (propertyMapping && propertyMapping.spoUserPropertyName) {
        // Get SharePoint User Profile Property Value
        let value = userProfile[propertyMapping.spoUserPropertyName];

        // If the User Property has a value and encryption is enabled, then set value to encrypted hash
        if (value && propertyMapping.encryptSHA256) {
          value = sha256(value).toString();
        }

        // If is a gtmPropertyName is specified, then use that property name in the GTM dataLayer
        if (propertyMapping.gtmPropertyName) {
          obj[propertyMapping.gtmPropertyName] = value;
        }
        // Else use the user profile property name in the GTM dataLayer
        else {
          obj[propertyMapping.spoUserPropertyName] = value;
        }
      }
      return obj;
    }, {});
    return dataLayerProperties;
  }

  private async getCurrentUserProperties(): Promise<any> {
    const { disableCache, cacheLifetimeMins } = this.getProperties();
    const cacheKey = 'GoogleAnalyticsExtensionUserProfileProperties';
    const cacheExpiration = 1000 * 60 * (undefined !== cacheLifetimeMins && !isNaN(cacheLifetimeMins) ? cacheLifetimeMins : DEFAULT_PROPERTIES.cacheLifetimeMins);
    const cachedUserProfile = this.cacheGet(cacheKey);

    if (!disableCache && cachedUserProfile) {
      return cachedUserProfile;
    }
    else {
     const freshUserProfile = await this.fetchCurrentUserProperties();
     return this.cacheSet(cacheKey, freshUserProfile, cacheExpiration);
    }
  }

  private async fetchCurrentUserProperties(): Promise<any> {
    try {
      const response = await this.context.spHttpClient.get("/_api/SP.UserProfiles.PeopleManager/GetMyProperties", SPHttpClient.configurations.v1, {
        headers: { Accept: "application/json" }
      });
      if (response.ok) {
        const userProfile: { UserProfileProperties: { Key: string; Value: string; }[] } = await response.json();
        const userProfileFlattenedAllProperties = userProfile.UserProfileProperties.reduce((allProperties, currentProperty) => {
          if (currentProperty) {
            allProperties[currentProperty.Key] = currentProperty.Value;
          }
          return allProperties;
        }, {});
        return userProfileFlattenedAllProperties;
      }
      else {
        throw new Error(`[${response.status}] Unable to fetch current user's profile. ${response.statusText}`);
      }
    }
    catch (error) {
      Log.error(LOG_SOURCE, error);
    }
  }

  /**
   * Helper function to get an item by key from local storage
   */
  private cacheGet = (key) => {
    try {
      const valueStr = localStorage.getItem(key);
      if (valueStr) {
        const val = JSON.parse(valueStr);
        if (val) {
          return !(val.expiration && Date.now() > val.expiration) ? val.payload : null;
        }
      }
      return null;
    }
    catch (error) {
      Log.warn(LOG_SOURCE, `Unable to get cached data from localStorage with key ${key}. Error: ${error}`);
      return null;
    }
  }

  /**
   * Helper function to set an item by key into local storage
   */
  private cacheSet = (key, payload, expiresIn) => {
    const { disableCache } = this.getProperties();
    if (disableCache) return payload;
    try {
      const nowTicks = Date.now();
      const expiration = (expiresIn && nowTicks + expiresIn) || null;
      const cache = { payload, expiration };
      localStorage.setItem(key, JSON.stringify(cache));
      return this.cacheGet(key);
    }
    catch (error) {
      Log.warn(LOG_SOURCE, `Unable to set cached data into localStorage with key ${key}. Error: ${error}`);
      return null;
    }
  }

}
