import { MailFolder, OutlookItem } from "@microsoft/microsoft-graph-types"
import { MSALAuthProvider } from "authProvider"
import MSGraphPlugin from "MSGraphPlugin"
import { SelectMailFolderModal } from "selectMailFolderModal"

import * as Eta from 'eta'
import { MSGraphMailFolderAccess } from "types"


export class MailHandler {
    plugin:MSGraphPlugin

    constructor(plugin:MSGraphPlugin) {
        this.plugin = plugin
    }

    getMailFolders = async (authProvider: MSALAuthProvider, limit=250):Promise<MailFolder[]> => {
		// this could be so much simpler, if we wouldn't support self-hosted exchange installations...
		// otherwise, we could just query https://graph.microsoft.com/v1.0/me/mailfolders/delta?$select=displayname
		// and would be done...
		const graphClient = this.plugin.getGraphClient(authProvider)

		const getChildFoldersRecursive = async (fs:Array<MailFolder>): Promise<MailFolder[]> => {
			let result:Array<MailFolder> = []
			
			for (const f of fs) {
				const request = graphClient
					.api(`/me/mailfolders/${f.id}/childFolders?includeHiddenFolders=true`)
					.select("displayName,childFolderCount")
					.top(limit)

				const childFolders = (await request.get()).value

				result = result.concat(childFolders)

				const remainingChildFolders = childFolders.filter((f:MailFolder) => f.childFolderCount > 0)
				const remainingGrandChildren = await getChildFoldersRecursive(remainingChildFolders)

				result = result.concat(remainingChildFolders).concat(remainingGrandChildren)
			}

			return result
		}

		let request = graphClient
			.api("/me/mailfolders?includeHiddenFolders=true")
			.select("displayName,childFolderCount")
			.top(limit)
		
		let root_folders = (await request.get()).value

		root_folders = root_folders.concat(await getChildFoldersRecursive(root_folders.filter((f:MailFolder) => f.childFolderCount > 0)))

		request = graphClient
			.api("/me/mailFolders/searchfolders/childFolders")
			.select("displayName,childFolderCount")
			.top(limit)

		root_folders = root_folders.concat((await request.get()).value)

		return root_folders
	}

	selectMailFolder = async (account:string) => {
		const selector = new SelectMailFolderModal(this.plugin.app, this.plugin)

		const folders = await this.getMailFolders(this.plugin.msalProviders[account])
		selector.setFolders(folders)
		selector.open()
	}

	getMailsForFolder = async (mf: MSGraphMailFolderAccess) => {
        const authProvider = this.plugin.msalProviders[mf.provider]

        const graphClient = this.plugin.getGraphClient(authProvider)

		let request = graphClient.api(`/me/mailFolders/${mf.id}/messages`)

		if (mf.query !== undefined)
			request = request.query(mf.query)

		if (mf.limit !== undefined)
			request = request.top(mf.limit)

        if (mf.onlyFlagged)
            request = request.filter('flag/flagStatus eq \'flagged\'')

		return (await request.get()).value
	}

    getMailsForAllFolders = async () => {
        const mails:Record<string, [OutlookItem]> = {}
        for (const mf of this.plugin.settings.mailFolders) {
            mails[mf.displayName] = await this.getMailsForFolder(mf)
        }

        return mails
    }

	formatMails = (mails:Record<string,[any]>, as_tasks=false):string => {
		let result = ""

        for (const folder in mails) {
            result += `# ${folder}\n\n`

            for (const m of mails[folder]) {
                result += Eta.render(as_tasks
                    ? this.plugin.settings.flaggedMailTemplate 
                    : this.plugin.settings.mailTemplate, m) + "\n\n"
            }
   
            result += "\n"
        }

		return result
	}

    formatMailsForAllFolders = async (as_tasks=false): Promise<string> => {
        return this.formatMails(await this.getMailsForAllFolders(), as_tasks)
    }
}