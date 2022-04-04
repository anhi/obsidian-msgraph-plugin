import { MailFolder } from "@microsoft/microsoft-graph-types";
import MSGraphPlugin from "MSGraphPlugin";
import { App, FuzzySuggestModal, Notice } from "obsidian";

export class SelectMailFolderModal extends FuzzySuggestModal<MailFolder> {
  app:App
  plugin:MSGraphPlugin
  folders:MailFolder[]
  
  constructor(app:App, plugin:MSGraphPlugin) {
    super(app);

    this.app = app
    this.plugin = plugin
  }

  setFolders = (folders:MailFolder[]) => {
    this.folders = folders
  }

  getItems(): MailFolder[] {
    return this.folders
  }

  getItemText(f: MailFolder): string {
    return f.displayName;
  }

  onChooseItem(f: MailFolder, evt: MouseEvent | KeyboardEvent) {
    new Notice(`Selected ${f.displayName}`);
  }
}