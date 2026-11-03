import { App, ButtonComponent, Notice, PluginSettingTab, SecretComponent, Setting } from 'obsidian'
import { MailFolder } from '@microsoft/microsoft-graph-types'

import type MSGraphPlugin from './MSGraphPlugin'
import { AccountType, MSGraphAccount, MSGraphMailFolderAccess } from './types'
import { SelectMailFolderModal } from './selectMailFolderModal'
import { errorMessage } from './util'

import { defaultEventTemplate } from "./defaultEventTemplate"
import { defaultMailTemplate } from "./defaultMailTemplate"
import { defaultFlaggedMailTemplate } from "./defaultFlaggedMailTemplate"

export interface MSGraphPluginSettings {
	accounts: Array<MSGraphAccount>,
	mailFolders: Array<MSGraphMailFolderAccess>,
	eventTemplate: string,
	mailTemplate: string,
	flaggedMailTemplate: string,
}

export const DEFAULT_SETTINGS: MSGraphPluginSettings = {
	accounts: [new MSGraphAccount()],
	mailFolders: [new MSGraphMailFolderAccess()],
	eventTemplate: defaultEventTemplate,
	mailTemplate: defaultMailTemplate,
	flaggedMailTemplate: defaultFlaggedMailTemplate,
}

export class MSGraphPluginSettingsTab extends PluginSettingTab {
	plugin: MSGraphPlugin;

	constructor(app: App, plugin: MSGraphPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const {containerEl} = this;

		containerEl.empty();

		this.add_account_setting()

		this.add_mail_settings()

		new Setting(this.containerEl).setName("Templates").setHeading()
		new Setting(this.containerEl).setDesc(
			"Eta.js templates for converting MSGraph objects to Markdown. Use <%= … %> for content coming from mails " +
			"and events: it escapes Markdown, so that a malicious mail or invitation cannot inject links, embeds, HTML " +
			"or code blocks (e.g. dataviewjs) into your notes. <%~ … %> inserts the value unchanged."
		)

		this.add_event_template()
		this.add_mail_template()
	}

	private setAccountEnabled = async (account: MSGraphAccount, enabled: boolean): Promise<void> => {
		if (enabled) {
			const name = account.displayName.trim()
			if (!name) {
				new Notice("MSGraph: please enter a display name first.")
			} else if (this.plugin.settings.accounts.some((a) => a !== account && a.displayName.trim() === name)) {
				new Notice(`MSGraph: there is already an account called "${name}".`)
			} else {
				account.displayName = name
				account.enabled = this.plugin.registerAccount(account) !== null
			}
		} else {
			account.enabled = false
			await this.plugin.unregisterAccount(account.displayName)
		}

		await this.plugin.saveSettings()
		this.display()
	}

	add_account_setting(): void {
		new Setting(this.containerEl).setName("MSGraph Access").setHeading()

		new Setting(this.containerEl).setDesc(
			"Each account you want to query MSGraph with needs to be authorized. " +
			"Disable an account to change its configuration. Client secrets of confidential applications are " +
			"kept in Obsidian's keychain, not in the plugin settings."
		)

		new Setting(this.containerEl)
			.setName("Add Account")
			.setDesc("Add MSGraph Account")
			.addButton((button: ButtonComponent) => {
				button
					.setTooltip("Add additional MSGraph account")
					.setButtonText("+")
					.setCta()
					.onClick(async () => {
						this.plugin.settings.accounts.push(new MSGraphAccount());
						await this.plugin.saveSettings();
						this.display();
					});
			});

		this.plugin.settings.accounts.forEach(
			(account, index) => {
				const locked = account.enabled

				const div = this.containerEl.createEl("div");
				div.addClass("msgraph_div");

				const title = this.containerEl.createEl("h4", {
					text: `Account #${index}` + (account.lastUsername ? ` (${account.lastUsername})` : ""),
				});
				title.addClass("msgraph_title");

				const s = new Setting(this.containerEl)
					.addExtraButton((extra) => {
						extra
							.setIcon("cross")
							.setTooltip("Delete")
							.onClick(async () => {
								if (account.enabled) {
									await this.plugin.unregisterAccount(account.displayName)
								}
								this.plugin.settings.accounts.splice(
									this.plugin.settings.accounts.indexOf(account),
									1
								);
								await this.plugin.saveSettings();
								// Force refresh
								this.display();
							});
					})
					.addText((text) => {
						const t = text
							.setPlaceholder("Display Name")
							.setValue(account.displayName)
							.setDisabled(locked)
							.onChange(async (new_value) => {
								account.displayName = new_value;
								await this.plugin.saveSettings();
							});
						t.inputEl.addClass("msgraph_display_name");

						return t;
					})
					.addText((text) => {
						const t = text
							.setPlaceholder("Client Id")
							.setValue(account.clientId)
							.setDisabled(locked)
							.onChange(async (new_value) => {
								account.clientId = new_value.trim();
								await this.plugin.saveSettings();
							});
						t.inputEl.addClass("msgraph_client_id");

						return t;
					})
					.addComponent((controlEl) => {
						// controlEl is shared by all controls of this row, so use a wrapper of our own
						const wrapper = controlEl.createDiv({ cls: "msgraph_client_secret" })
						wrapper.setAttribute("aria-label", "Client secret (optional, from keychain)")
						return new SecretComponent(this.app, wrapper)
							.setValue(account.clientSecretId)
							.setDisabled(locked)
							.onChange(async (new_value) => {
								account.clientSecretId = new_value ?? "";
								await this.plugin.saveSettings();
							})
					})
					.addText((text) => {
						const t = text
							.setPlaceholder("authority")
							.setValue(account.authority)
							.setDisabled(locked)
							.onChange(async (new_value) => {
								account.authority = new_value.trim();
								await this.plugin.saveSettings();
							});
						t.inputEl.addClass("msgraph_authority");

						return t;
					})
					.addToggle((toggle) => {
						toggle.setValue(account.enabled)
							.setTooltip(this.plugin.checkProvider(account.displayName)
								? "Authorized – disable to sign out"
								: "Authorize account")
							.onChange((new_value) => this.setAccountEnabled(account, new_value))
					})
					.addDropdown((dropdown) => {
						dropdown.addOptions(
							{MSGraph: "MSGraph", EWS: "EWS"}
						)
						.setValue(account.type)
						.setDisabled(locked)
						.onChange(async (new_value) => {
							account.type = new_value as AccountType
							await this.plugin.saveSettings()
						})

						dropdown.selectEl.addClass("msgraph_provider_type")
					})
					.addText((text) => {
						const t = text
							.setPlaceholder("EWS URL (https://…)")
							.setValue(account.baseUri)
							.setDisabled(locked)
							.onChange(async (new_value) => {
								account.baseUri = new_value.trim();
								await this.plugin.saveSettings();
							});
						t.inputEl.addClass("msgraph_baseuri");

						return t;
					})

				s.infoEl.remove();

				div.appendChild(title);
				div.appendChild(s.settingEl);
			}
		);
	}

	add_mail_settings = (): void => {
		new Setting(this.containerEl).setName("Mail related settings").setHeading()

		new Setting(this.containerEl).setDesc("Mail folders you want to access");

		new Setting(this.containerEl)
			.setName("Add Mail Folder")
			.addButton((button: ButtonComponent) => {
				button
					.setTooltip("Add additional mail folder")
					.setButtonText("+")
					.setCta()
					.onClick(async () => {
						this.plugin.settings.mailFolders.push(new MSGraphMailFolderAccess());
						await this.plugin.saveSettings();
						this.display();
					});
			});

		this.plugin.settings.mailFolders.forEach(
			(folder) => {
				const div = this.containerEl.createEl("div");
				div.addClass("msgraph_div");

				const s = new Setting(this.containerEl)
					.addExtraButton((extra) => {
						extra
							.setIcon("cross")
							.setTooltip("Delete")
							.onClick(async () => {
								this.plugin.settings.mailFolders.splice(
									this.plugin.settings.mailFolders.indexOf(folder),
									1
								);
								await this.plugin.saveSettings();
								// Force refresh
								this.display();
							});
					})
				.addText((text) => {
					const t = text
						.setPlaceholder("Display Name")
						.setValue(folder.displayName)
						.onChange(async (new_value) => {
							folder.displayName = new_value;
							await this.plugin.saveSettings();
						});
					t.inputEl.addClass("msgraph_display_name");

					return t;
				})
				.addDropdown((dropdown) => {
					const options: Record<string, string> = {invalid: "Choose Provider"}

					for (const account of this.plugin.settings.accounts) {
						if (account.displayName.trim()) {
							options[account.displayName] = account.displayName
						}
					}
					const d = dropdown
						.addOptions(options)
						.setValue(folder.provider in options ? folder.provider : "invalid")
						.onChange(async (new_value) => {
							folder.provider = new_value
							await this.plugin.saveSettings()
						});
					d.selectEl.addClass("msgraph_input")
				})
				.addText((text) => {
					text
						.setPlaceholder("<Folder Id>")
						.setValue(folder.id)
						.onChange(async (new_value) => {
							folder.id = new_value.trim()
							await this.plugin.saveSettings()
						});
				})
				.addButton((button) => {
					button
						.setButtonText("Browse")
						.setTooltip("Select mail folder")
						.onClick(async () => {
							const provider = this.plugin.msalProviders.get(folder.provider)
							if (!provider) {
								new Notice("MSGraph: please choose an enabled account first.")
								return
							}

							const modal = new SelectMailFolderModal(this.app, this.plugin)
							modal.onChooseItem = async (f: MailFolder) => {
								folder.id = f.id ?? ""
								await this.plugin.saveSettings()
								this.display()
							}

							try {
								modal.setFolders(await this.plugin.mailHandler.getMailFolders(provider))
								modal.open()
							} catch (e) {
								new Notice(`MSGraph: could not list mail folders: ${errorMessage(e)}`)
								console.error(e)
							}
						})
				})
				.addText((text) => {
					const t = text
						.setPlaceholder("Maximum number of mails to return")
						.setValue(folder.limit.toString())
						.onChange(async (new_value) => {
							const new_limit = parseInt(new_value, 10)
							if (!isNaN(new_limit) && new_limit > 0) {
								folder.limit = Math.min(new_limit, 1000)
								await this.plugin.saveSettings()
							}
						})
					t.inputEl.addClass("msgraph_number")
					t.inputEl.type = "number"
					t.inputEl.min = "1"
					t.inputEl.max = "1000"
				})
				.addDropdown((dropdown) => {
					dropdown
						.addOptions({all: "All emails", flagged: "Only flagged emails"})
						.setValue(folder.onlyFlagged ? 'flagged' : 'all')
						.onChange(async (new_value) => {
							folder.onlyFlagged = new_value === 'flagged'
							await this.plugin.saveSettings()
						})

					dropdown.selectEl.addClass("msgraph_input")
				})

				s.infoEl.remove();

				div.appendChild(s.settingEl);
			});
	}

	private add_template_setting = (
		name: string,
		desc: string,
		defaultTemplate: string,
		key: 'eventTemplate' | 'mailTemplate' | 'flaggedMailTemplate'
	): void => {
		new Setting(this.containerEl)
			.setName(name)
			.setDesc(desc)
			.addButton((button) => {
				button
					.setButtonText("Default")
					.setTooltip("Restore default template")
					.onClick(async () => {
						this.plugin.settings[key] = defaultTemplate
						await this.plugin.saveSettings();
						// Force refresh
						this.display();
					});
			})
			.addTextArea((text) => {
				const t = text
					.setPlaceholder("Template Text")
					.setValue(this.plugin.settings[key])
					.onChange(async (new_value) => {
						this.plugin.settings[key] = new_value
						await this.plugin.saveSettings()
					})
				t.inputEl.addClass("msgraph_template")
			})
	}

	add_event_template = (): void => {
		this.add_template_setting("Event Template", "Template for rendering Events", defaultEventTemplate, 'eventTemplate')
	}

	add_mail_template = (): void => {
		this.add_template_setting("Mail Template", "Template for rendering Mail items", defaultMailTemplate, 'mailTemplate')
		this.add_template_setting("Flagged mail template", "Template for rendering flagged mail items as tasks",
			defaultFlaggedMailTemplate, 'flaggedMailTemplate')
	}
}
