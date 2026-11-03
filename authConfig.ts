import type { MSGraphAccount } from './types'

export const redirectUri = 'obsidian://msgraph'
export const msgraphScopes = ["offline_access", "User.read", "Calendars.read", "Mail.read"]
export const ewsScopes = (baseUri: string) => [`${baseUri}/EWS.AccessAsUser.All`, "offline_access"]

/**
 * Validates the EWS base URI. The OAuth access token is sent to this host, so only https is allowed.
 * Returns the URI without trailing slashes.
 */
export const normalizeEwsBaseUri = (baseUri: string): string => {
    let url: URL
    try {
        url = new URL(baseUri.trim())
    } catch {
        throw new Error(`Invalid EWS URL "${baseUri}"`)
    }
    if (url.protocol !== 'https:') {
        throw new Error(`EWS URL must use https:// (got "${baseUri}")`)
    }
    return url.origin + url.pathname.replace(/\/+$/, '')
}

/** Validates the authority (token issuer) URL. */
export const validateAuthority = (authority: string): string => {
    let url: URL
    try {
        url = new URL(authority.trim())
    } catch {
        throw new Error(`Invalid authority "${authority}"`)
    }
    if (url.protocol !== 'https:') {
        throw new Error(`Authority must use https:// (got "${authority}")`)
    }
    return url.href.replace(/\/+$/, '')
}

export const getScopes = (account: MSGraphAccount): string[] =>
    account.type === "EWS" ? ewsScopes(normalizeEwsBaseUri(account.baseUri)) : msgraphScopes
