import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

export type CachedToken = {
    displayName: string,
    accessToken: string,
}

export class MSGraphAccount {
    displayName  = ""
    clientId     = ""
    clientSecret = ""
    authority    = "https://login.microsoftonline.com/common"
    enabled      = false
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