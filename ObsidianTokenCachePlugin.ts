import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node"
import { App, Notice } from "obsidian"

/** Key prefix used by versions <= 0.9.9 in the global (not vault-specific) localStorage. */
const LEGACY_KEY_PREFIX = 'msal-tokencache-'
/** Key prefix for the vault-specific localStorage (App.saveLocalStorage). */
const KEY_PREFIX = 'msgraph-tokencache-'

interface SafeStorage {
    isEncryptionAvailable(): boolean
    encryptString(plainText: string): Buffer
    decryptString(encrypted: Buffer): string
}

/**
 * Electron's safeStorage (OS keychain backed encryption), as exposed by Obsidian.
 * Returns null if encryption is not available on this system.
 */
export const getSafeStorage = (): SafeStorage | null => {
    try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const safeStorage: SafeStorage | undefined = (window as any).electron?.remote?.safeStorage
        if (safeStorage && safeStorage.isEncryptionAvailable()) {
            return safeStorage
        }
    } catch (e) {
        console.error("MSGraph: could not access Electron safeStorage", e)
    }
    return null
}

let warnedAboutMissingEncryption = false

/**
 * MSAL cache plugin that stores the token cache encrypted (Electron safeStorage) in the
 * vault-specific localStorage. If encryption is unavailable, tokens are only kept in memory
 * and never written to disk unencrypted.
 */
export class ObsidianTokenCachePlugin implements ICachePlugin {
    private readonly key: string
    private memoryCache: string | null = null

    constructor(private readonly app: App, public readonly displayName: string) {
        this.key = KEY_PREFIX + displayName
        this.migrateLegacyCache()
    }

    /**
     * Older versions stored the cache in the localStorage shared by *all* vaults.
     * Move it into the vault-specific storage and delete the shared copy.
     */
    private migrateLegacyCache(): void {
        const legacyKey = LEGACY_KEY_PREFIX + this.displayName
        let legacy: string | null = null
        try {
            legacy = window.localStorage.getItem(legacyKey)
        } catch {
            return
        }
        if (legacy === null) return

        const safeStorage = getSafeStorage()
        if (safeStorage && this.app.loadLocalStorage(this.key) === null) {
            try {
                const plain = safeStorage.decryptString(Buffer.from(legacy, 'latin1'))
                this.persist(plain)
            } catch (e) {
                console.warn("MSGraph: could not migrate legacy token cache, a new login will be required", e)
            }
        }
        window.localStorage.removeItem(legacyKey)
    }

    private loadPersisted(): string | null {
        const stored = this.app.loadLocalStorage(this.key)
        if (typeof stored !== 'string') return null

        const safeStorage = getSafeStorage()
        if (!safeStorage) return null

        try {
            return safeStorage.decryptString(Buffer.from(stored, 'base64'))
        } catch (e) {
            console.error("MSGraph: could not decrypt token cache, discarding it", e)
            this.app.saveLocalStorage(this.key, null)
            return null
        }
    }

    private persist(serialized: string): void {
        const safeStorage = getSafeStorage()
        if (!safeStorage) {
            if (!warnedAboutMissingEncryption) {
                warnedAboutMissingEncryption = true
                new Notice("MSGraph: secure storage is not available on this system. " +
                    "Login tokens are kept in memory only, you will have to sign in again after a restart.")
            }
            return
        }
        this.app.saveLocalStorage(this.key, safeStorage.encryptString(serialized).toString('base64'))
    }

    public async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
        if (this.memoryCache === null) {
            this.memoryCache = this.loadPersisted()
        }
        if (this.memoryCache) {
            cacheContext.tokenCache.deserialize(this.memoryCache)
        }
    }

    public async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
        if (cacheContext.cacheHasChanged) {
            this.memoryCache = cacheContext.tokenCache.serialize()
            this.persist(this.memoryCache)
        }
    }

    /** Delete the token cache from memory and storage. */
    public deleteFromCache(): void {
        this.memoryCache = null
        this.app.saveLocalStorage(this.key, null)
    }

    /** Is there a (possibly expired) token cache for this account? */
    public isInitialized(): boolean {
        return this.memoryCache !== null || this.app.loadLocalStorage(this.key) !== null
    }
}
