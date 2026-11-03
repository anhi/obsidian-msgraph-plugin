## Microsoft Graph Plugin

This plugin connects Obsidian (https://obsidian.md) to MS Graph, the central gateway to access and modify
information stored in MS 365.

Currently, the plugin allows to access to calendar items and emails.

### Requirements

Obsidian 1.11.4 or newer (desktop only).

### Installation

Download the release zip-file and extract it inside the .obsidian/plugins - folder in your Vault. Then, activate the MSGraph
plugin in the community plugins - section of the settings.

### Configuration

To connect to MS Graph, the plugin needs to authenticate with an app registered via Microsoft Identity Platform. To register, either
follow the instructions at https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app, or use the
Azure command line interface (https://docs.microsoft.com/en-us/cli/azure/install-azure-cli). In any case, the app needs to be configured
as a Single Page Application with the redirect URI "obsidian://msgraph".

#### Using Azure portal

In App registrations, click on your application, then Manage => Authentication => Platform configurations => Add a platform => Single-page application.
Set the redirect URI to obsidian://msgraph. Write down the Application (client) ID in the App overview.

If you want to use a confidential application, navigate to your application, then

Certificates & secrets => Client secrets => New client secret.

If the sign-in fails with error `AADSTS9002327` ("Tokens issued for the 'Single-Page Application' client-type may only
be redeemed via cross-origin requests"), add the redirect URI `obsidian://msgraph` under the platform
"Mobile and desktop applications" instead.

#### Using the Azure CLI

Create the app with

```az ad app create --display-name obsidian-graph-client --reply-urls obsidian://msgraph/auth```

Optionally, you can add a client secret for a confidential application through the ```---password``` flag.

Write down the appId. Then, configure the app as an SPA with

```az rest --method PATCH     --uri 'https://graph.microsoft.com/v1.0/applications/{appID}'    --headers 'Content-Type=application/json'  --body "{spa:{redirectUris:['obsidian://msgraph']}}"```

where you replace {appId} with the value you wrote down in the last step.

### Configuring obsidian

Once the application has been registered, you can configure the plugin in Obsidian. Activate the plugin and navigate to its settings. For each account you want to query, add an MSGraph account.
The display name can be chosen freely, the Client ID corresponds to the appId you wrote down earlier. If you configured a password, store it in
Obsidian's keychain (Settings => Keychain) and select it as the client secret of the account. The secret is never written to the plugin's
`data.json`; secrets stored there by older versions of this plugin are moved to the keychain automatically. The default for the authority field
should work in most cases. Activate the account by switching the toggle to the right. To change the configuration of an account, deactivate it first
(this also signs it out).

Login tokens are stored encrypted (using the operating system's keychain) and are specific to the vault.

If you want to access your email, configure Mail Folders in the following section of the settings dialog. Finally, if you want to fine-tune the rendering of events, mails, or tasks, you can change the
Eta.js - templates in the settings.

#### Templates and security

Mails and calendar invitations can be sent to you by anyone. If their content were inserted into your notes unchanged, a malicious mail
could, for example, add a `dataviewjs` code block, a Templater command, HTML, or a remote image (tracking pixel) to your note. Therefore:

  - `<%= value %>` escapes Markdown. Use it for everything that comes from a mail or an event (subject, body, sender, ...).
  - `<%~ value %>` inserts the value unchanged. Only use it for values you produced yourself or already escaped with `it.escapeMarkdown(...)`.
  - `it.safeUrl(url)` returns the URL only if it is an http(s) URL, for use in Markdown links.

The default templates follow these rules. If you customized a template with an older version of the plugin, please review it.

To use the plugin, open a note, open the Obsidian command palette (CTRL+p or CMD+p), and select one of the MSGraph - functions. Your browser will show a login window and ask for permission to access your data.


### Usage

Currently, the plugin allows to

  - access all events for today from all calendars and write them to the current note
  - access all mails or all flagged mails inside the mail folders specified in the settings and write them to the current note, optionally formatted
    as tasks

### Example usage

With the help of Templater (https://github.com/SilentVoid13/Templater), it is simple to add a list of today's events and flagged mails to the daily note. Simply add something like
this

```
## Meetings

<% await this.app.plugins.plugins['obsidian-msgraph-plugin'].calendarHandler.formatEventsForToday() %>


### Overdue Emails

<% await this.app.plugins.plugins['obsidian-msgraph-plugin'].mailHandler.formatOverdueMailsForAllFolders() %>
```

to your daily note template.

