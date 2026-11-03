import {
    AccountInfo,
    AuthError,
    AuthenticationResult,
    ConfidentialClientApplication,
    CryptoProvider,
    PublicClientApplication,
} from '@azure/msal-node'
import { AuthenticationProvider, Client } from '@microsoft/microsoft-graph-client'
import { App, Notice, ObsidianProtocolData } from 'obsidian'

import { getScopes, normalizeEwsBaseUri, redirectUri, validateAuthority } from './authConfig'
import { ObsidianTokenCachePlugin } from './ObsidianTokenCachePlugin'
import type { MSGraphAccount } from './types'

/** How long to wait for the user to complete the login in the browser. */
const LOGIN_TIMEOUT_MS = 5 * 60 * 1000

/** Errors where an interactive login would not help (we are offline, DNS fails, ...). */
const NETWORK_ERROR_CODES = new Set(['network_error', 'no_network_connectivity', 'endpoints_resolution_error'])

interface PendingLogin {
    state: string
    verifier: string
    resolve: (result: AuthenticationResult) => void
    reject: (error: Error) => void
}

export class MSALAuthProvider implements AuthenticationProvider {
    readonly msalClient: PublicClientApplication | ConfidentialClientApplication
    readonly cachePlugin: ObsidianTokenCachePlugin
    readonly graphClient: Client

    private readonly crypto = new CryptoProvider()
    private readonly scopes: string[]
    private pendingLogin: PendingLogin | null = null
    private loginPromise: Promise<string> | null = null

    /**
     * @throws Error if the account configuration is invalid
     */
    constructor(app: App, readonly account: MSGraphAccount, clientSecret: string | null) {
        if (!account.displayName.trim()) throw new Error("Account has no display name")
        if (!account.clientId.trim()) throw new Error("Account has no client ID")
        if (account.type === "EWS") normalizeEwsBaseUri(account.baseUri)

        this.scopes = getScopes(account)
        this.cachePlugin = new ObsidianTokenCachePlugin(app, account.displayName)

        const auth = {
            clientId:  account.clientId.trim(),
            authority: validateAuthority(account.authority),
        }
        const cache = { cachePlugin: this.cachePlugin }

        this.msalClient = clientSecret
            ? new ConfidentialClientApplication({ auth: { ...auth, clientSecret }, cache })
            : new PublicClientApplication({ auth, cache })

        this.graphClient = Client.initWithMiddleware({ authProvider: this })
    }

    isInitialized = (): boolean => {
        return this.cachePlugin.isInitialized()
    }

    /**
     * Returns an access token from the cache (refreshing it if necessary), or null if an
     * interactive login is required.
     * @throws on network errors
     */
    getTokenSilently = async (): Promise<string | null> => {
        const accounts = await this.msalClient.getTokenCache().getAllAccounts()
        if (accounts.length === 0) return null

        try {
            const result = await this.msalClient.acquireTokenSilent({ account: accounts[0], scopes: this.scopes })
            return result?.accessToken || null
        } catch (error) {
            if (error instanceof AuthError && NETWORK_ERROR_CODES.has(error.errorCode)) {
                throw error
            }
            console.warn(`MSGraph: silent token acquisition for "${this.account.displayName}" failed`,
                error instanceof AuthError ? error.errorCode : error)
            return null
        }
    }

    /**
     * Called by the Graph client before every request. Resolves to an access token or rejects.
     */
    getAccessToken = async (): Promise<string> => {
        const token = await this.getTokenSilently()
        return token ?? this.login()
    }

    /** Starts an interactive login. Concurrent callers share the same login. */
    login = (): Promise<string> => {
        if (!this.loginPromise) {
            this.loginPromise = this.interactiveLogin().finally(() => {
                this.pendingLogin = null
                this.loginPromise = null
            })
        }
        return this.loginPromise
    }

    private async interactiveLogin(): Promise<string> {
        // fresh PKCE codes and an unguessable state for every login attempt
        const { verifier, challenge } = await this.crypto.generatePkceCodes()
        const state = this.crypto.createNewGuid()

        const authCodeUrl = await this.msalClient.getAuthCodeUrl({
            scopes: this.scopes,
            redirectUri,
            prompt: "select_account",
            state,
            codeChallenge: challenge,
            codeChallengeMethod: "S256",
        })

        let timer: number | undefined
        const result = new Promise<AuthenticationResult>((resolve, reject) => {
            this.pendingLogin = { state, verifier, resolve, reject }
            timer = window.setTimeout(
                () => reject(new Error(`Login for "${this.account.displayName}" timed out`)),
                LOGIN_TIMEOUT_MS
            )
        })

        new Notice(`MSGraph: please sign in to "${this.account.displayName}" in your browser.`)
        // Obsidian opens external URLs in the system browser
        window.open(authCodeUrl)

        try {
            const response = await result
            await this.removeOtherAccounts(response.account)
            return response.accessToken
        } finally {
            window.clearTimeout(timer)
        }
    }

    /**
     * Handles the obsidian://msgraph redirect. Returns true if the redirect belongs to a login
     * started by this provider.
     */
    handleRedirect = (query: ObsidianProtocolData): boolean => {
        const pending = this.pendingLogin
        if (!pending || typeof query.state !== 'string' || query.state !== pending.state) {
            return false
        }
        // the state is single-use
        this.pendingLogin = null

        if (query.error) {
            pending.reject(new Error(`Sign-in for "${this.account.displayName}" failed: ${query.error_description || query.error}`))
            return true
        }
        if (typeof query.code !== 'string' || !query.code) {
            pending.reject(new Error(`Sign-in for "${this.account.displayName}" failed: no authorization code received`))
            return true
        }

        this.msalClient.acquireTokenByCode({
            code: query.code,
            codeVerifier: pending.verifier,
            redirectUri,
            scopes: this.scopes,
            clientInfo: typeof query.client_info === 'string' ? query.client_info : undefined,
        }).then(pending.resolve, pending.reject)

        return true
    }

    /** Only keep the account that just signed in, so that getTokenSilently() picks the right one. */
    private async removeOtherAccounts(keep: AccountInfo | null): Promise<void> {
        if (!keep) return
        const cache = this.msalClient.getTokenCache()
        for (const account of await cache.getAllAccounts()) {
            if (account.homeAccountId !== keep.homeAccountId) {
                await cache.removeAccount(account)
            }
        }
    }

    /** Cancel a pending login. */
    dispose = (): void => {
        this.pendingLogin?.reject(new Error("Login cancelled"))
        this.pendingLogin = null
    }

    /** Cancel a pending login and remove all cached tokens of this account. */
    signOut = async (): Promise<void> => {
        this.dispose()
        try {
            const cache = this.msalClient.getTokenCache()
            for (const account of await cache.getAllAccounts()) {
                await cache.removeAccount(account)
            }
        } finally {
            this.cachePlugin.deleteFromCache()
        }
    }
}

export async function refreshAllTokens(providers: Iterable<MSALAuthProvider>): Promise<void> {
    await Promise.allSettled([...providers]
        .filter((provider) => provider.account.enabled && provider.isInitialized())
        .map((provider) => provider.getTokenSilently()))
}
