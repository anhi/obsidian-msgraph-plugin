import { TokenCacheContext, ICachePlugin } from "@azure/msal-common";

const { safeStorage } = require('@electron/remote')

export const MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY_PREFIX = 'msal-tokencache-'

export class ObsidianTokenCachePlugin implements ICachePlugin {

    public displayName: string
    private MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY: string
    public acquired: boolean

    constructor(displayName: string) {
        this.displayName = displayName
        this.MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY = MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY_PREFIX + this.displayName
        this.acquired = false
    }

    /**
     * Reads from local storage and decrypts. We don't care about efficiency here, otherwise
     * we could cache values in memory. But then we would have to prevent race conditions between
     * successive method calls, and this seems excessive for our use case.
     */
    public async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
        //console.info("Executing before cache access");

        const encryptedCache: string = localStorage.getItem(this.MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY)
        const cache = (encryptedCache !== null)
                ? safeStorage.decryptString(Buffer.from(encryptedCache, 'latin1'))
                : ""

        cacheContext.tokenCache.deserialize(cache);
    }

    /**
     * Encrypts and writes to local storage.
     */
    public async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
        //console.info("Executing after cache access");
        
        if (cacheContext.cacheHasChanged) {
            const serializedAccounts = cacheContext.tokenCache.serialize()
            
            localStorage.setItem(
                this.MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY,
                safeStorage.encryptString(serializedAccounts).toString('latin1')
            )
        }
    }

    /** Delete the token cache from local storage.
     */
    public async deleteFromCache(): Promise<void> {
        localStorage.removeItem(this.MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY)
        this.acquired = false
    }

    /** Has this cache been properly initialized? */
    public isInitialized(): boolean {
        return (localStorage.getItem(this.MSAL_ACCESS_TOKEN_LOCALSTORAGE_KEY) !== null)
    }
}