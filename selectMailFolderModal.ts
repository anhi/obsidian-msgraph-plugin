import { MailFolder } from "@microsoft/microsoft-graph-types";
import { App, FuzzySuggestModal, Notice } from "obsidian";

import type MSGraphPlugin from "./MSGraphPlugin";

export class SelectMailFolderModal extends FuzzySuggestModal<MailFolder> {
  plugin: MSGraphPlugin
  folders: MailFolder[] = []

  constructor(app: App, plugin: MSGraphPlugin) {
    super(app);

    this.plugin = plugin
    this.setPlaceholder("Select a mail folder")
  }

  setFolders = (folders: MailFolder[]) => {
    this.folders = folders
  }

  getItems(): MailFolder[] {
    return this.folders
  }

  getItemText(f: MailFolder): string {
    return f.displayName ?? f.id ?? "";
  }

  onChooseItem(f: MailFolder, evt: MouseEvent | KeyboardEvent) {
    new Notice(`Selected ${f.displayName}`);
  }
}
