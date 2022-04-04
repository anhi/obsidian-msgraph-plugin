import { ClientApplication } from "@azure/msal-node";
//import { FilePersistenceWithDataProtection, DataProtectionScope } from "@azure/msal-node-extensions";
import { requestConfig } from 'authConfig'

import { shell } from "electron";
import { PublicClientApplication } from '@azure/msal-node';
import { AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import { CryptoProvider } from '@azure/msal-node';
import { AuthorizationCodeRequest } from '@azure/msal-node';
import { SilentFlowRequest } from "@azure/msal-node";
import { AuthenticationResult } from "@azure/msal-node";
import { PkceCodes } from "@azure/msal-common";
import { MSGraphAccount } from "types";
import MSGraphPlugin from "MSGraphPlugin";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { ObsidianTokenCachePlugin } from "ObsidianTokenCachePlugin";


const { safeStorage } = require('@electron/remote')

export const MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY = 'msal-access_token'
export const MSGRAPH_ACCOUNTS_LOCALSTORAGE_KEY = 'msgraph-accounts'

export class MSALAuthProvider implements AuthenticationProvider {
    // todo: somehow persist the tokens
    authConfig = {
        verifier: "",
        challenge: "",
    }
    msalClient: ClientApplication = null
    account: MSGraphAccount = null
    
    cachePlugin: ObsidianTokenCachePlugin

    constructor(account: MSGraphAccount) { 
        this.account = account

        this.cachePlugin = new ObsidianTokenCachePlugin(account.displayName)
        if (account.clientSecret && account.clientSecret.trim()) {
            this.msalClient = new ConfidentialClientApplication({
                auth: {
                    clientId:      account.clientId,
                    clientSecret:  account.clientSecret,
                    authority:     account.authority,
                },
                cache: {
                    cachePlugin: this.cachePlugin
                }
            });
        } else {
            this.msalClient = new PublicClientApplication({
                auth: {
                    clientId:  account.clientId,
                    authority: account.authority,
                },
                cache: {
                    cachePlugin: this.cachePlugin
                }
            });
        }

        const cryptoProvider = new CryptoProvider();
        cryptoProvider.generatePkceCodes()
            .then((codes: PkceCodes) => {
                this.authConfig.challenge = codes.challenge
                this.authConfig.verifier  = codes.verifier
            })
    } 

    removeAccessToken = async () => {
        this.cachePlugin.deleteFromCache()
    }

    isInitialized = (): boolean => {
        return this.cachePlugin.isInitialized()
    }

    getTokenSilently = async (): Promise<string> => {
        // retrieve all cached accounts
        const accounts = await this.msalClient.getTokenCache().getAllAccounts();

        if (accounts.length > 0) {
            // todo: logic to choose the correct account
            //       for now, just use the first one
            const account = accounts[0]
            const silentRequest: SilentFlowRequest = {
                ...requestConfig.request.silentRequest,
                account: account
            }

            return this.msalClient.acquireTokenSilent(silentRequest)
                .then((authResponse: AuthenticationResult) => {
                    return authResponse.accessToken
                })
                .then((accessToken: string) => {
                    console.info("Successfully obtained access token from cache!")
                    return accessToken
                })
                .catch((error: any)  => {
                    return ""
            })
        } else {
            return ""
        }
    }

	/**
	 * This method will get called before every request to the msgraph server
	 * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
	 * Basically this method will contain the implementation for getting and refreshing accessTokens
	 */
    getAccessToken = async () => {
        const access_token = await this.getTokenSilently()
        
        if (access_token !== "") {
            return access_token
        } else  {
            msalLogin(this)

            let total_waiting_time = 0
            const max_waiting_time = 60000 // 1 minute
            const ms = 500

            while (this.cachePlugin.acquired == false && total_waiting_time <= max_waiting_time) {
                await new Promise(resolve => {
                    setTimeout(resolve, ms)
                })
                total_waiting_time += ms
            }

            if (this.cachePlugin.acquired == false) {
                console.log("Could not acquire token!")
                return ""
            } else {
                console.info("Successfully logged in!")
                return await this.getTokenSilently()
            }
        }
    }
}

export function msalLogin(msalProvider: MSALAuthProvider) {
    const pkceCodes = {
        challengeMethod: "S256", // Use SHA256 Algorithm
        verifier: msalProvider.authConfig.verifier,
        challenge: msalProvider.authConfig.challenge
      };
  
    const authCodeUrlParams = { 
        ...requestConfig.request.authCodeUrlParameters, // redirectUri, scopes
        state: msalProvider.account.displayName,
        codeChallenge: pkceCodes.challenge, // PKCE Code Challenge
        codeChallengeMethod: pkceCodes.challengeMethod, // PKCE Code Challenge Method
    };
  
  
    msalProvider.msalClient.getAuthCodeUrl(authCodeUrlParams)
        .then((response:any) => {
            shell.openExternal(response);
        })
        .catch((error:any) => console.log(JSON.stringify(error)));
}

export function msalRedirect(plugin: MSGraphPlugin, query: any) {
    const displayName = query.state;

    if (!(displayName in plugin.msalProviders)) {
        console.error("Invalid auth request: unknown account!")
        return
    }

    const authProvider = plugin.msalProviders[displayName]

    // Add PKCE code verifier to token request object
    const tokenRequest: AuthorizationCodeRequest = {
        ...requestConfig.request.tokenRequest,
        code: query.code as string,
        codeVerifier: authProvider.authConfig.verifier, // PKCE Code Verifier
        clientInfo: query.client_info as string
    };

    authProvider.msalClient.acquireTokenByCode(tokenRequest).then((response: AuthenticationResult) => {
        authProvider.cachePlugin.acquired = true
    }).catch((error: any) => {
        console.log(error)
        authProvider.cachePlugin.acquired = false
    })
}