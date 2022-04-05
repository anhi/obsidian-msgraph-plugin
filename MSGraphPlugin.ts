import { Editor, MarkdownView, ObsidianProtocolData, Plugin } from 'obsidian';

import { MSALAuthProvider, msalRedirect } from 'authProvider';
import { Client } from '@microsoft/microsoft-graph-client';


import { MSGraphPluginSettings, DEFAULT_SETTINGS, MSGraphPluginSettingsTab } from 'msgraphPluginSettings';
import { MSGraphAccount } from 'types';

import { MailHandler } from 'mailHandler'


import { CalendarHandler } from 'calendarHandler';

export default class MSGraphPlugin extends Plugin {
	settings: MSGraphPluginSettings;
	msalProviders: Record<string, MSALAuthProvider> = {}
	graphClient: Client

	calendarHandler: CalendarHandler
	mailHandler: MailHandler

	authenticateAccount = (account: MSGraphAccount) => {
		const provider = new MSALAuthProvider(account)

		this.msalProviders[account.displayName] = provider

		return provider
	}

	getGraphClient = (msalProvider: MSALAuthProvider): Client => {
		return Client.initWithMiddleware({
			debugLogging: true,
			authProvider: msalProvider
		});
	}

	checkProvider = (displayName: string) => {
		return (
			displayName in this.msalProviders && 
			this.msalProviders[displayName].isInitialized()
		)
	}

	async onload() {
		await this.loadSettings();

		this.calendarHandler = new CalendarHandler(this)
		this.mailHandler     = new MailHandler(this)

		this.registerObsidianProtocolHandler('msgraph', (query: ObsidianProtocolData) => {msalRedirect(this, query);})

		// register all stored providers
		for (const account of this.settings.accounts) {
			if (account.displayName.trim() !== "" && account.enabled) {
				this.authenticateAccount(account)
			}
		}
	
		// Append today's events, sorted by start time
		this.addCommand({
			id: 'append-todays-events-by-start',
			name: 'Append today\'s Events, sorted by start date',
			editorCallback: async (editor: Editor, view: MarkdownView) => {
				const result = this.calendarHandler.formatEvents(await this.calendarHandler.getEventsForToday())

				editor.replaceSelection(result);
			}
		});
	
		this.addCommand({
			id: 'get-mails-from-all-folders',
			name: 'Append mails from all folders registered in the settings',
			editorCallback: async (editor: Editor, view: MarkdownView) => {
				const result = await this.mailHandler.formatMails(await this.mailHandler.getMailsForAllFolders(), false)

				editor.replaceSelection(result)
			}
		})

		this.addCommand({
			id: 'get-mails-from-all-folders-as-tasks',
			name: 'Append mails from all folders registered in the settings and format as tasks',
			editorCallback: async (editor: Editor, view: MarkdownView) => {
				const result = await this.mailHandler.formatMails(await this.mailHandler.getMailsForAllFolders(), true)

				editor.replaceSelection(result)
			}
		})


		// This adds a settings tab so the user can configure various aspects of the plugin
		this.addSettingTab(new MSGraphPluginSettingsTab(this.app, this));
	}

	onunload() {

	}

	async loadSettings() {
		this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}