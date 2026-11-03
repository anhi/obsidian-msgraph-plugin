import { MailFolder, Message } from "@microsoft/microsoft-graph-types"
import { moment, Notice } from "obsidian"

import type { MSALAuthProvider } from "./authProvider"
import type MSGraphPlugin from "./MSGraphPlugin"
import { SelectMailFolderModal } from "./selectMailFolderModal"
import { eta, templateHelpers } from "./templating"
import { MSGraphMailFolderAccess } from "./types"
import { errorMessage, getAllPages } from "./util"

/** Graph allows at most 1000 messages per page. */
const MAX_MAILS = 1000

export class MailHandler {
	plugin: MSGraphPlugin

	constructor(plugin: MSGraphPlugin) {
		this.plugin = plugin
	}

	getMailFolders = async (authProvider: MSALAuthProvider, limit=250): Promise<MailFolder[]> => {
		// this could be so much simpler, if we wouldn't support self-hosted exchange installations...
		// otherwise, we could just query https://graph.microsoft.com/v1.0/me/mailfolders/delta?$select=displayname
		// and would be done...
		const graphClient = authProvider.graphClient

		const listFolders = (path: string) => getAllPages<MailFolder>(
			graphClient,
			graphClient.api(path).query("includeHiddenFolders=true").select("displayName,childFolderCount").top(limit),
			10000
		)

		const getChildFoldersRecursive = async (parents: MailFolder[]): Promise<MailFolder[]> => {
			let result: MailFolder[] = []

			for (const parent of parents) {
				if (!parent.id) continue
				const childFolders = await listFolders(`/me/mailFolders/${encodeURIComponent(parent.id)}/childFolders`)
				const grandChildren = await getChildFoldersRecursive(childFolders.filter((f) => (f.childFolderCount ?? 0) > 0))

				result = result.concat(childFolders, grandChildren)
			}

			return result
		}

		const rootFolders = await listFolders("/me/mailFolders")
		let folders = rootFolders.concat(await getChildFoldersRecursive(rootFolders.filter((f) => (f.childFolderCount ?? 0) > 0)))

		try {
			folders = folders.concat(await listFolders("/me/mailFolders/searchfolders/childFolders"))
		} catch (e) {
			// not every mailbox has a search folder root
			console.info("MSGraph: could not list search folders", e)
		}

		return folders
	}

	selectMailFolder = async (account: string) => {
		const authProvider = this.plugin.msalProviders.get(account)
		if (!authProvider) throw new Error(`Account "${account}" is not enabled`)

		const selector = new SelectMailFolderModal(this.plugin.app, this.plugin)
		selector.setFolders(await this.getMailFolders(authProvider))
		selector.open()
	}

	getMailsForFolder = async (mf: MSGraphMailFolderAccess): Promise<Message[]> => {
		const authProvider = this.plugin.msalProviders.get(mf.provider)
		if (!authProvider) throw new Error(`account "${mf.provider}" is not enabled`)

		const folderId = mf.id.trim()
		if (!folderId) throw new Error("no mail folder selected")

		const graphClient = authProvider.graphClient
		let request = graphClient.api(`/me/mailFolders/${encodeURIComponent(folderId)}/messages`)

		if (mf.query?.trim())
			request = request.query(mf.query.trim())

		const limit = (Number.isInteger(mf.limit) && mf.limit > 0) ? Math.min(mf.limit, MAX_MAILS) : 100
		request = request.top(limit)

		if (mf.onlyFlagged)
			request = request.filter('flag/flagStatus eq \'flagged\'')

		return getAllPages<Message>(graphClient, request, limit)
	}

	/**
	 * Fetches the mails of all configured folders. Folders that fail are reported and skipped,
	 * folders without account or folder id are ignored.
	 */
	getMailsForAllFolders = async (): Promise<Record<string, Message[]>> => {
		const mails: Record<string, Message[]> = {}
		for (const mf of this.plugin.settings.mailFolders) {
			if (!mf.provider || mf.provider === "invalid" || !mf.id.trim()) continue

			const name = mf.displayName || mf.id
			try {
				mails[name] = (mails[name] ?? []).concat(await this.getMailsForFolder(mf))
			} catch (e) {
				new Notice(`MSGraph: could not fetch mails for "${name}": ${errorMessage(e)}`)
				console.error(e)
			}
		}

		return mails
	}

	formatMails = (mails: Record<string, unknown[]>, as_tasks=false): string => {
		let result = ""
		const template = as_tasks ? this.plugin.settings.flaggedMailTemplate : this.plugin.settings.mailTemplate

		try {
			for (const folder in mails) {
				result += `# ${folder}\n\n`

				for (const m of mails[folder]) {
					result += eta.renderString(template, {...(m as object), ...templateHelpers}) + "\n\n"
				}

				result += "\n"
			}
		} catch (e) {
			new Notice(`MSGraph: error in mail template: ${errorMessage(e)}`)
			console.error(e)
		}

		return result
	}

	formatMailsForAllFolders = async (as_tasks=false): Promise<string> => {
		return this.formatMails(await this.getMailsForAllFolders(), as_tasks)
	}

	dueFilter = (m: Message): boolean => {
		const due = m?.flag?.dueDateTime?.dateTime
		return due ? moment.utc(due).valueOf() <= Date.now() : true
	}

	formatOverdueMailsForAllFolders = async (): Promise<string> => {
		const ms = await this.getMailsForAllFolders()

		for (const folder in ms) {
			ms[folder] = ms[folder].filter(this.dueFilter)
		}

		return this.formatMails(ms, true)
	}
}
