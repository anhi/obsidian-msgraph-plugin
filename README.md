## Microsoft Graph Plugin

This plugin connects Obsidian (https://obsidian.md) to MS Graph, the central gateway to access and modify
information stored in MS 365.

Currently, the plugin allows to access to calendar items and emails.

### Installation

Download the release zip-file and extract it inside the .obsidian/plugins - folder in your Vault. Then, activate the MSGraph
plugin in the community plugins - section of the settings.

### Configuration


### Usage

Currently, the plugin allows to

  - access all events for today from all calendars and write them to the current note
  - access all mails or all flaggeed mails inside the mail folders specified in the settings and write them to the current note, optionally formatted
    as tasks

#### Example usage

With the help of Templater (https://github.com/SilentVoid13/Templater), it is simple to add a list of today's events and flagged mails to the daily note, use  to automatically add 


### API Documentation

See https://github.com/obsidianmd/obsidian-api
