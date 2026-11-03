import { App, Modal, Setting } from 'obsidian'

export interface LoginInfo {
    /** Display name of the account in the plugin settings */
    displayName: string
    /** What is being accessed, e.g. "Microsoft 365 (Microsoft Graph)" */
    service: string
    /** User that was signed in last time, if known */
    username?: string
    /** Why a new sign-in is necessary, if known */
    reason?: string
}

/**
 * Tells the user which account has to be signed in before the browser is opened, and stays
 * open until the sign-in is completed or cancelled.
 */
export class LoginModal extends Modal {
    private settled = false
    private actionsEl: HTMLElement

    constructor(
        app: App,
        private readonly info: LoginInfo,
        private readonly onOpenBrowser: () => void,
        private readonly onCancel: () => void,
    ) {
        super(app)
    }

    onOpen(): void {
        const { contentEl, info } = this
        this.titleEl.setText(`Sign in to "${info.displayName}"`)

        contentEl.createEl('p', { text: `MSGraph needs you to sign in to the account "${info.displayName}" (${info.service}).` })
        if (info.username) {
            const p = contentEl.createEl('p', { text: 'Sign in as: ' })
            p.createEl('strong', { text: info.username })
        } else {
            contentEl.createEl('p', { text: 'Make sure to choose the right Microsoft account on the sign-in page.' })
        }
        if (info.reason) {
            contentEl.createEl('p', { text: `Reason: ${info.reason}`, cls: 'msgraph_login_reason' })
        }

        this.actionsEl = contentEl.createDiv()
        new Setting(this.actionsEl)
            .addButton((b) => b.setButtonText('Cancel').onClick(() => this.cancel()))
            .addButton((b) => b.setButtonText('Open sign-in page').setCta().onClick(() => this.openBrowser()))
    }

    openBrowser(): void {
        this.onOpenBrowser()
        this.actionsEl.empty()
        new Setting(this.actionsEl)
            .setDesc('Waiting for the sign-in in your browser…')
            .addButton((b) => b.setButtonText('Cancel').onClick(() => this.cancel()))
            .addButton((b) => b.setButtonText('Open again').onClick(() => this.onOpenBrowser()))
    }

    cancel(): void {
        this.close()
    }

    /** Close the dialog without cancelling the login (it has completed or failed). */
    finish(): void {
        this.settled = true
        this.close()
    }

    onClose(): void {
        this.contentEl.empty()
        if (!this.settled) {
            this.settled = true
            this.onCancel()
        }
    }
}
