import { AuthError, INetworkModule, NetworkRequestOptions, NetworkResponse } from '@azure/msal-node'
import { requestUrl } from 'obsidian'

import { errorMessage } from './util'

/**
 * Network client for MSAL based on Obsidian's requestUrl.
 *
 * MSAL Node's default client uses the global fetch(), which inside Obsidian is the browser's fetch:
 * it is subject to CORS and sends an `Origin: app://obsidian.md` header, which makes Microsoft Entra ID
 * treat the token requests as single-page-application requests. requestUrl is executed by Obsidian
 * outside of the browser context (like Node's https module used by earlier MSAL versions).
 */
export class ObsidianNetworkClient implements INetworkModule {
    sendGetRequestAsync<T>(url: string, options?: NetworkRequestOptions, timeout?: number): Promise<NetworkResponse<T>> {
        return this.send<T>(url, 'GET', options, timeout)
    }

    sendPostRequestAsync<T>(url: string, options?: NetworkRequestOptions): Promise<NetworkResponse<T>> {
        return this.send<T>(url, 'POST', options)
    }

    private async send<T>(url: string, method: 'GET' | 'POST', options?: NetworkRequestOptions, timeout?: number): Promise<NetworkResponse<T>> {
        const request = requestUrl({
            url,
            method,
            headers: options?.headers,
            body: method === 'POST' ? (options?.body ?? '') : undefined,
            throw: false,
        })

        let timer: number | undefined
        let response
        try {
            response = timeout
                ? await Promise.race([
                    request,
                    new Promise<never>((_, reject) => {
                        timer = window.setTimeout(() => reject(new Error('Request timeout')), timeout)
                    }),
                ])
                : await request
        } catch (e) {
            throw new AuthError('network_error', `Network request failed: ${errorMessage(e)}`)
        } finally {
            window.clearTimeout(timer)
        }

        let body: T
        try {
            body = JSON.parse(response.text) as T
        } catch {
            throw new AuthError('token_parsing_error', `Failed to parse response from ${new URL(url).host} (HTTP ${response.status})`)
        }

        const headers: Record<string, string> = {}
        for (const [name, value] of Object.entries(response.headers ?? {})) {
            headers[name.toLowerCase()] = String(value)
        }

        return { headers, body, status: response.status }
    }
}
