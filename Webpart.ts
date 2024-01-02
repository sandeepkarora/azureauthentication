import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { getSP } from './pnpjsConfig';
// import { sp } from "@pnp/sp/presets/all";
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import { IReadonlyTheme } from '@microsoft/sp-component-base';


import { msalConfig } from "./Config";
import { PublicClientApplication } from "@azure/msal-browser";

export interface ISampleWebPartProps {
  description: string;
  context: WebPartContext;
}

// import { IPublicClientApplication } from "@azure/msal-browser";
// type AppProps = {
//   pca: IPublicClientApplication
// };

const msalInstance = new PublicClientApplication(msalConfig);

export default class SampleWebPart extends BaseClientSideWebPart<ISampleWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  constructor() {
    super();

    let Current_User = sessionStorage.getItem("current_user_details");
    let user = (Current_User != "" && Current_User != null) ? JSON.parse(Current_User) : "";        //v2.0:replaced || with &&
    //console.log(Current_User, user);
    if (user == "" || user == null) {
      let loginRespoinse = this.login().then(() => {
        msalInstance.handleRedirectPromise().
          then((loginResponse) => {
            // console.log("Token Acquired.. " + loginResponse);
            sessionStorage.setItem("current_user_details", JSON.stringify(loginResponse));
            if (JSON.stringify(loginResponse) != undefined) {
              this.render()
            }
          }
          ).catch((error) => {
            console.log(error);
          });

      }

      );
    }
    else {
      //alert("User is already logged in..");
    }
  }

  public async login(): Promise<any> {
    const loginRequest:any = {
      scopes: ["Scope"] // optional Array<string>
     //scopes: [process.env.SPFX_scopes] // optional Array<string>
    };
    let loginResponse:any = null;
    try {
      loginResponse = await msalInstance.loginRedirect(loginRequest);
    } catch (err) {
      // handle error
      console.log('err login', err);
    }
    return loginResponse;
  }

  public render(): void {
    const element: React.ReactElement<IAceguiProps> = React.createElement(
      Acegui,
      {
        isLoading:true,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );
   
    ReactDom.render(element, this.domElement);
  }

  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
  // }
  protected async onInit(): Promise<void> {
    await this._getEnvironmentMessage().then(message => {
          this._environmentMessage = message;
        });

    await super.onInit();

    //Initialize our _sp object that we can then use in other packages without having to pass around the context.
    //  Check out pnpjsConfig.ts for an example of a project setup file.
    getSP(this.context);
    // return super.onInit().then(_ => {
    //   sp.setup({
    //     spfxContext: this.context
    //   });
    // });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
