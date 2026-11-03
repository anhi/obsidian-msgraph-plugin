import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

export type AccountType = "MSGraph" | "EWS"

export class MSGraphAccount {
    displayName    = ""
    clientId       = ""
    /**
     * Name of the secret in Obsidian's secret storage (Settings → Keychain) that holds the
     * client secret for confidential applications. The secret itself is never stored in data.json.
     */
    clientSecretId = ""
    /**
     * @deprecated Plain-text client secret from versions <= 0.9.9. Only read once to migrate it into
     * Obsidian's secret storage, never written back.
     */
    clientSecret?: string
    authority      = "https://login.microsoftonline.com/common"
    enabled        = false
    type: AccountType = "MSGraph"
    baseUri        = ""
}

export type EventWithProvider = MicrosoftGraph.Event & {provider: string}

export class MSGraphMailFolderAccess {
    displayName = ""
    provider    = ""
    id          = ""
    limit       = 100
    query       = ""
    onlyFlagged = false
}
