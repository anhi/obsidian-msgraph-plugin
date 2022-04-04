import { LogLevel } from '@azure/msal-node';

export const redirectUri = 'obsidian://msgraph'
export const scopes = ["offline_access", "User.read", "Calendars.read", "Mail.read"]

export const requestConfig = {
    request:
    {
        authCodeUrlParameters: {
            scopes: scopes,
            redirectUri: redirectUri,
            prompt: "select_account",
        },
        tokenRequest: {
            code: "",
            redirectUri: redirectUri,
            scopes: scopes,
        },
        silentRequest: {
            scopes: scopes,
        }
    },
};