
import { App, ButtonComponent, Setting } from 'obsidian'
import { PluginSettingTab } from 'obsidian'
import MSGraphPlugin from './MSGraphPlugin'

import { MSGraphAccount, MSGraphMailFolderAccess } from 'types';
import { SelectMailFolderModal } from 'selectMailFolderModal';
import { MailFolder } from '@microsoft/microsoft-graph-types';

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

		containerEl.createEl('h2', {text: 'Settings for MSGraphPlugin.'});

		this.add_account_setting()

		this.add_mail_settings()

		this.containerEl.createEl("h2", { text: "Templates" });
		const descHeading = document.createDocumentFragment();
        descHeading.append(
			"Templates for converting MSGraph objects to Markdown."
        );
        new Setting(this.containerEl).setDesc(descHeading);

		this.add_event_template()
		this.add_mail_template()
	}

	add_account_setting(): void {
        this.containerEl.createEl("h2", { text: "MSGraph Access" });

        const descHeading = document.createDocumentFragment();
        descHeading.append(
			"Each account you want to query MSGraph with needs to be authorized."
        );

        new Setting(this.containerEl).setDesc(descHeading);

        new Setting(this.containerEl)
            .setName("Add Account")
            .setDesc("Add MSGraph Account")
            .addButton((button: ButtonComponent) => {
                button
                    .setTooltip("Add additional MSGraph account")
                    .setButtonText("+")
                    .setCta()
                    .onClick(() => {
                        this.plugin.settings.accounts.push(new MSGraphAccount());
                        this.plugin.saveSettings();
                        this.display();
                    });
            });

        this.plugin.settings.accounts.forEach(
            (account, index) => {
				const div = this.containerEl.createEl("div");
                div.addClass("msgraph_div");

                const title = this.containerEl.createEl("h4", {
                    text: "Account #" + index,
                });
                title.addClass("msgraph_title");

                const s = new Setting(this.containerEl)
				.addExtraButton((extra) => {
					extra
						.setIcon("cross")
						.setTooltip("Delete")
						.onClick(() => {
							this.plugin.settings.accounts.splice(
								index,
								1
							);
							this.plugin.saveSettings();
							// Force refresh
							this.display();
						});
				})
				.addText((text) => {
					const t = text
						.setPlaceholder("Display Name")
						.setValue(account.displayName)
						.onChange((new_value) => {
							this.plugin.settings.accounts[index].displayName = new_value;
							this.plugin.saveSettings();
						});
					t.inputEl.addClass("msgraph_display_name");

					return t;
				})
				.addText((text) => {
					const t = text
						.setPlaceholder("Client Id")
						.setValue(account.clientId)
						.onChange((new_value) => {
							this.plugin.settings.accounts[index].clientId = new_value;
							this.plugin.saveSettings();
						});
					t.inputEl.addClass("msgraph_client_id");

					return t;
				})
				.addText((text) => {
					const t = text
						.setPlaceholder("Client Secret (optional)")
						.setValue(account.clientSecret)
						.onChange((new_value) => {
							this.plugin.settings.accounts[index].clientSecret = new_value;
							this.plugin.saveSettings();
						});
					t.inputEl.addClass("msgraph_client_secret");

					return t;
				})
				.addText((text) => {
					const t = text
						.setPlaceholder("authority")
						.setValue(account.authority)
						.onChange((new_value) => {
							this.plugin.settings.accounts[index].authority = new_value;
							this.plugin.saveSettings();
						});
					t.inputEl.addClass("msgraph_authority");

					return t;
				})
				.addToggle((toggle) => {
					toggle.setValue(this.plugin.checkProvider(this.plugin.settings.accounts[index].displayName))
						.setTooltip("Authorize account")
						.onChange((new_value) => {
							const account = this.plugin.settings.accounts[index]

							if (new_value) {
								this.plugin.authenticateAccount(account)
								account.enabled = true
							} else {
								// todo: de-authorize token
								this.plugin.msalProviders[account.displayName].removeAccessToken()
								delete this.plugin.msalProviders[account.displayName]
								account.enabled = false
							}
						})
				})
                
				s.infoEl.remove();
				
                div.appendChild(title);
                div.appendChild(this.containerEl.lastChild);
            }
        );
    }

	add_mail_settings = (): void => {
		this.containerEl.createEl("h2", { text: "Mail related settings" });

        const descHeading = document.createDocumentFragment();
        descHeading.append(
			"Mail folders you want to access"
        );

        new Setting(this.containerEl).setDesc(descHeading);

        new Setting(this.containerEl)
            .setName("Add Mail Folder")
            .addButton((button: ButtonComponent) => {
                button
                    .setTooltip("Add additional mail folder")
                    .setButtonText("+")
                    .setCta()
                    .onClick(() => {
                        this.plugin.settings.mailFolders.push(new MSGraphMailFolderAccess());
                        this.plugin.saveSettings();
                        this.display();
                    });
            });

        this.plugin.settings.mailFolders.forEach(
            (folder, index) => {
				const div = this.containerEl.createEl("div");
                div.addClass("msgraph_div");

				const s = new Setting(this.containerEl)
					.addExtraButton((extra) => {
						extra
							.setIcon("cross")
							.setTooltip("Delete")
							.onClick(() => {
								this.plugin.settings.mailFolders.splice(
									index,
									1
								);
								this.plugin.saveSettings();
								// Force refresh
								this.display();
							});
					})
				.addText((text) => {
					const t = text
						.setPlaceholder("Display Name")
						.setValue(folder.displayName)
						.onChange((new_value) => {
							folder.displayName = new_value;
							this.plugin.saveSettings();
						});
					t.inputEl.addClass("msgraph_display_name");

					return t;
				})
				.addDropdown((dropdown) => {
					const options: Record<string, string> = {invalid: "Choose Provider"}
	
					for (const account of this.plugin.settings.accounts) {
						options[account.displayName] = account.displayName
					}
					const d = dropdown
						.addOptions(options)
						.setValue(folder.provider)
						.onChange((new_value) => {
							folder.provider = new_value
							this.plugin.saveSettings()
							// Force refresh
							this.display();
						});
						d.selectEl.addClass("msgraph_input")
				})
				.addText((text) => {
					const t = text
						.setPlaceholder("<Id or path>")
						.setValue(folder.id)
						.onChange((new_value) => {
							folder.id = new_value
							this.plugin.saveSettings()
							// Force refresh
							this.display();
						});
				})
				.addButton((button) => {
					button
						.setButtonText("Browse")
						.setTooltip("Select search folder")
						.onClick(() => {
							if (this.plugin.checkProvider(folder.provider)) {
								const modal = new SelectMailFolderModal(this.app, this.plugin)
								modal.onChooseItem = (f:MailFolder, evt:Event) => {
									folder.id = f.id
									this.plugin.saveSettings()
									this.display()
								}
	
								this.plugin.mailHandler
									.getMailFolders(this.plugin.msalProviders[folder.provider])
									.then((folders) => modal.setFolders(folders))
									.then(() => modal.open())
	
							}
						})
				})
				.addText((text) => {
					const t = text
						.setPlaceholder("Maximum number of mails to return")
						.setValue(folder.limit.toString())
						.onChange((new_value) => {
							const new_limit = parseInt(new_value, 10)
							if (!isNaN(new_limit)) {
								folder.limit = new_limit
								this.plugin.saveSettings()
							} else {
								this.display()
							}
						})
						t.inputEl.addClass("msgraph-number")
				})
				.addDropdown((dropdown) => {
					dropdown
						.addOptions({all: "All emails", flagged: "Only flagged emails"})
						.setValue(folder.onlyFlagged ? 'flagged' : 'all')
						.onChange((new_value) => {
							folder.onlyFlagged = new_value === 'flagged'
							this.plugin.saveSettings()
						})
						
						dropdown.selectEl.addClass("msgraph_input")
					}
				)

				s.infoEl.remove();
				
                div.appendChild(this.containerEl.lastChild);
			});
			
	}

	add_event_template = (): void => {
        new Setting(this.containerEl)
            .setName("Event Template")
            .setDesc("Template for rendering Events")
			.addButton((button) => {
				button
					.setButtonText("Default")
					.setTooltip("Restore default template")
					.onClick(() => {
						this.plugin.settings.eventTemplate = defaultEventTemplate
						this.plugin.saveSettings();
						// Force refresh
						this.display();
					});
			})
			.addTextArea((text) => {
				const t = text
					.setPlaceholder("Template Text")
					.setValue(this.plugin.settings.eventTemplate)
					.onChange((new_value) => {
						this.plugin.settings.eventTemplate = new_value
						this.plugin.saveSettings()
					})
				t.inputEl.addClass("msgraph_template")
			})
		}
	
	add_mail_template = (): void => {
		new Setting(this.containerEl)
			.setName("Mail Template")
			.setDesc("Template for rendering Mail items")
			.addButton((button) => {
				button
					.setButtonText("Default")
					.setTooltip("Restore default template")
					.onClick(() => {
						this.plugin.settings.mailTemplate = defaultMailTemplate
						this.plugin.saveSettings();
						// Force refresh
						this.display();
					});
			})
			.addTextArea((text) => {
				const t = text
					.setPlaceholder("Template Text")
					.setValue(this.plugin.settings.mailTemplate)
					.onChange((new_value) => {
						this.plugin.settings.mailTemplate = new_value
						this.plugin.saveSettings()
					})
				t.inputEl.addClass("msgraph_template")
			})

		new Setting(this.containerEl)
			.setName("Flagged mail template")
			.setDesc("Template for rendering flagged mail items as tasks")
			.addButton((button) => {
				button
					.setButtonText("Default")
					.setTooltip("Restore default template")
					.onClick(() => {
						this.plugin.settings.flaggedMailTemplate = defaultFlaggedMailTemplate
						this.plugin.saveSettings();
						// Force refresh
						this.display();
					});
			})
			.addTextArea((text) => {
				const t = text
					.setPlaceholder("Template Text")
					.setValue(this.plugin.settings.flaggedMailTemplate)
					.onChange((new_value) => {
						this.plugin.settings.flaggedMailTemplate = new_value
						this.plugin.saveSettings()
					})
				t.inputEl.addClass("msgraph_template")
			})

	}
        
}
