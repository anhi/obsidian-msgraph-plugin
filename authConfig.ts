import { LogLevel } from '@azure/msal-node';

export const redirectUri = 'obsidian://msgraph'
export const msgraph_scopes = ["offline_access", "User.read", "Calendars.read", "Mail.read"]
export const ews_scopes = (baseUri:string) => [`${baseUri}/EWS.AccessAsUser.All`, "offline_access"]

export const requestConfig = {
    request:
    {
        authCodeUrlParameters: {
            scopes: msgraph_scopes,
            redirectUri: redirectUri,
            prompt: "select_account",
        },
        tokenRequest: {
            code: "",
            redirectUri: redirectUri,
            scopes: msgraph_scopes,
        },
        silentRequest: {
            scopes: msgraph_scopes,
        }
    },
};

export const ewsRequestConfig = (baseUri:string) => {
    return {
        request:
        {
            authCodeUrlParameters: {
                scopes: ews_scopes(baseUri),
                redirectUri: redirectUri,
                prompt: "select_account",
            },
            tokenRequest: {
                code: "",
                redirectUri: redirectUri,
                scopes: ews_scopes(baseUri),
            },
            silentRequest: {
                scopes: ews_scopes(baseUri),
            }
        },
    }
}