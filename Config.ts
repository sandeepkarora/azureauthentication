import { PublicClientApplication, PopupRequest } from "@azure/msal-browser";

export const msalConfig : any = {
    auth: {
        // 'Application (client) ID' of app registration in Azure portal - this value is a GUID
        clientId:  'Application (client) ID',
        //process.env.SPFX_clientId,
        // Full directory URL, in the form of https://login.microsoftonline.com/<tenant>
        authority: "https://login.microsoftonline.com/<tenantId>",
        //process.env.SPFX_authority,
        // Full redirect URL, in form of http://localhost:3000
        redirectUri: "Full redirect URL, in form of http://localhost:3000 or sharepoint site url",
        //process.env.SPFX_redirectUri,
        //redirectUri: window.location.href,
        //  scopes: ["api://UIb5u74f-b05g-4dib-u2e1-911138056990/Access_API"]
        scopes: ["Scope"
            //process.env.SPFX_scopes
        ]
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {
        loggerOptions: {
            loggerCallback: (level: any, message: any, containsPii: any) => {
                if (containsPii) {
                    return;
                }
                switch (level) {
                    // case msal.LogLevel.Error:		
                    //     console.error(message);		
                    //     return;		
                    // case msal.LogLevel.Info:		
                    //     console.info(message);		
                    //     return;		
                    // case msal.LogLevel.Verbose:		
                    //     console.debug(message);		
                    //     return;		
                    // case msal.LogLevel.Warning:		
                    //     console.warn(message);		
                    //     return;		
                }
            }
        }
    }
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */

export const msalInstance: PublicClientApplication = new PublicClientApplication(
    msalConfig 
);
export const loginRequest: PopupRequest = {
    scopes: ["User.Read"]
};

/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const tokenRequest = {
    scopes: ["Access_API"],
    forceRefresh: false // Set this to "true" to skip a cached token and go to the server to get a new token
};

// Add here the endpoints for MS Graph API services you would like to use.
export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me"
};