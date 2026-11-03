import { Editor, Notice, ObsidianProtocolData, Plugin } from 'obsidian'

import { MSALAuthProvider, refreshAllTokens } from './authProvider'
import { CalendarHandler } from './calendarHandler'
import { defaultEventTemplate } from './defaultEventTemplate'
import { defaultFlaggedMailTemplate } from './defaultFlaggedMailTemplate'
import { defaultMailTemplate } from './defaultMailTemplate'
import {
	legacyDefaultEventTemplate,
	legacyDefaultFlaggedMailTemplate,
	legacyDefaultMailTemplate,
} from './legacyTemplates'
import { MailHandler } from './mailHandler'
import { ObsidianTokenCachePlugin } from './ObsidianTokenCachePlugin'
import { DEFAULT_SETTINGS, MSGraphPluginSettings, MSGraphPluginSettingsTab } from './msgraphPluginSettings'
import { MSGraphAccount, MSGraphMailFolderAccess } from './types'
import { errorMessage } from './util'

export default class MSGraphPlugin extends Plugin {
	settings: MSGraphPluginSettings
	/** Registered (enabled) accounts, keyed by display name. */
	msalProviders = new Map<string, MSALAuthProvider>()

	calendarHandler: CalendarHandler
	mailHandler: MailHandler

	/**
	 * Creates and registers the auth provider for an account.
	 * Returns null (and shows a notice) if the account configuration is invalid.
	 */
	registerAccount = (account: MSGraphAccount): MSALAuthProvider | null => {
		try {
			let clientSecret: string | null = null
			if (account.clientSecretId) {
				clientSecret = this.app.secretStorage.getSecret(account.clientSecretId)
				if (!clientSecret) {
					throw new Error(`Secret "${account.clientSecretId}" not found in the keychain`)
				}
			}

			this.msalProviders.get(account.displayName)?.dispose()

			const provider = new MSALAuthProvider(this.app, account, clientSecret)
			this.msalProviders.set(account.displayName, provider)
			return provider
		} catch (e) {
			new Notice(`MSGraph: cannot enable account "${account.displayName}": ${errorMessage(e)}`)
			console.error(e)
			return null
		}
	}

	/** Unregisters an account and deletes its cached tokens. */
	unregisterAccount = async (displayName: string): Promise<void> => {
		const provider = this.msalProviders.get(displayName)
		this.msalProviders.delete(displayName)
		if (provider) {
			await provider.signOut()
		} else {
			// the account may never have been registered (e.g. invalid configuration), still remove stored tokens
			new ObsidianTokenCachePlugin(this.app, displayName).deleteFromCache()
		}
	}

	checkProvider = (displayName: string): boolean => {
		return this.msalProviders.get(displayName)?.isInitialized() ?? false
	}

	private handleAuthRedirect = (query: ObsidianProtocolData) => {
		for (const provider of this.msalProviders.values()) {
			if (provider.handleRedirect(query)) return
		}
		console.warn("MSGraph: ignoring authentication redirect that does not belong to a pending login")
	}

	private insertIntoEditor = async (editor: Editor, producer: () => Promise<string>) => {
		try {
			editor.replaceSelection(await producer())
		} catch (e) {
			new Notice(`MSGraph: ${errorMessage(e)}`)
			console.error(e)
		}
	}

	async onload() {
		await this.loadSettings()

		this.calendarHandler = new CalendarHandler(this)
		this.mailHandler     = new MailHandler(this)

		this.registerObsidianProtocolHandler('msgraph', this.handleAuthRedirect)

		// register all stored providers
		for (const account of this.settings.accounts) {
			if (account.displayName.trim() !== "" && account.enabled) {
				this.registerAccount(account)
			}
		}

		// Append today's events, sorted by start time
		this.addCommand({
			id: 'append-todays-events-by-start',
			name: 'Append today\'s Events, sorted by start date',
			editorCallback: (editor: Editor) => this.insertIntoEditor(editor,
				async () => this.calendarHandler.formatEvents(await this.calendarHandler.getEventsForToday()))
		})

		this.addCommand({
			id: 'get-mails-from-all-folders',
			name: 'Append mails from all folders registered in the settings',
			editorCallback: (editor: Editor) => this.insertIntoEditor(editor,
				() => this.mailHandler.formatMailsForAllFolders(false))
		})

		this.addCommand({
			id: 'get-mails-from-all-folders-as-tasks',
			name: 'Append mails from all folders registered in the settings and format as tasks',
			editorCallback: (editor: Editor) => this.insertIntoEditor(editor,
				() => this.mailHandler.formatMailsForAllFolders(true))
		})

		this.addSettingTab(new MSGraphPluginSettingsTab(this.app, this))

		// Try to refresh the tokens every 30 minutes
		this.registerInterval(
			window.setInterval(() => { void refreshAllTokens(this.msalProviders.values()) }, 1000 * 60 * 30)
		)
	}

	onunload() {
		for (const provider of this.msalProviders.values()) {
			provider.dispose()
		}
		this.msalProviders.clear()
	}

	async loadSettings() {
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		const data: any = (await this.loadData()) ?? {}
		let changed = false

		const accounts: MSGraphAccount[] = Array.isArray(data.accounts)
			? data.accounts.map((a: Partial<MSGraphAccount>) => Object.assign(new MSGraphAccount(), a))
			: DEFAULT_SETTINGS.accounts.map(() => new MSGraphAccount())

		for (const account of accounts) {
			if (account.type !== "EWS") account.type = "MSGraph"
		}

		const mailFolders: MSGraphMailFolderAccess[] = Array.isArray(data.mailFolders)
			? data.mailFolders.map((f: Partial<MSGraphMailFolderAccess>) => Object.assign(new MSGraphMailFolderAccess(), f))
			: DEFAULT_SETTINGS.mailFolders.map(() => new MSGraphMailFolderAccess())

		const template = (value: unknown, legacyDefault: string, newDefault: string): string => {
			if (typeof value !== 'string') return newDefault
			// unmodified default templates from older versions are replaced by the escaping ones
			if (value === legacyDefault) {
				changed = true
				return newDefault
			}
			return value
		}

		this.settings = {
			accounts,
			mailFolders,
			eventTemplate:       template(data.eventTemplate, legacyDefaultEventTemplate, defaultEventTemplate),
			mailTemplate:        template(data.mailTemplate, legacyDefaultMailTemplate, defaultMailTemplate),
			flaggedMailTemplate: template(data.flaggedMailTemplate, legacyDefaultFlaggedMailTemplate, defaultFlaggedMailTemplate),
		}

		if (this.migrateClientSecrets()) changed = true

		if (changed) await this.saveSettings()
	}

	/**
	 * Older versions stored client secrets in plain text in data.json (which is part of the vault
	 * and often synced). Move them into Obsidian's secret storage.
	 */
	private migrateClientSecrets(): boolean {
		let changed = false
		this.settings.accounts.forEach((account, index) => {
			if (!('clientSecret' in account)) return

			const secret = account.clientSecret?.trim()
			if (secret) {
				try {
					account.clientSecretId = this.storeClientSecret(account, index, secret)
					new Notice(`MSGraph: the client secret of "${account.displayName}" was moved to Obsidian's keychain.`)
				} catch (e) {
					new Notice(`MSGraph: could not move the client secret of "${account.displayName}" to the keychain. ` +
						`Please enter it again in the settings.`)
					console.error(e)
				}
			}
			delete account.clientSecret
			changed = true
		})
		return changed
	}

	private storeClientSecret(account: MSGraphAccount, index: number, secret: string): string {
		const storage = this.app.secretStorage
		const slug = account.displayName.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
		const base = `msgraph-${slug || `account-${index + 1}`}-client-secret`

		let id = base
		for (let n = 2; storage.getSecret(id) !== null && storage.getSecret(id) !== secret; n++) {
			id = `${base}-${n}`
		}
		storage.setSecret(id, secret)
		return id
	}

	async saveSettings() {
		await this.saveData(this.settings)
	}
}
