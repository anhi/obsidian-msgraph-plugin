import MSGraphPlugin from "MSGraphPlugin";
import { htmlToMarkdown } from "obsidian";
import { EventWithProvider } from "types";

import { Event } from '@microsoft/microsoft-graph-types'

import luxon, { DateTime } from "luxon";

import * as Eta from 'eta'

export class CalendarHandler {
    plugin:MSGraphPlugin

    constructor(plugin:MSGraphPlugin) {
        this.plugin = plugin
    }

    getEventsForTimeRange = async (start:DateTime, end:DateTime):Promise<Record<string, [Event]>> => {
		const events_by_provider:Record<string, [Event]> = {}

		// todo: fetch in parallel using Promise.all()?
		for (const authProviderName in this.plugin.msalProviders) {
			const authProvider = this.plugin.msalProviders[authProviderName]
			const graphClient = this.plugin.getGraphClient(authProvider)

			const dateQuery = `startDateTime=${start.toISODate()}&endDateTime=${end.toISODate()}`;
			
			const events = await graphClient
				.api('/me/calendarView').query(dateQuery)
				//.select('subject,start,end')
				.orderby(`start/DateTime`)
				.get();

			events_by_provider[authProviderName] = events.value
		}

		return events_by_provider
	}

	getEventsForToday = async () => {
		const today = DateTime.now()
		const tomorrow = DateTime.now().plus({days: 1})
		
		return await this.getEventsForTimeRange(today, tomorrow)
	}

	flattenAndSortEventsByStartTime = (events_by_provider:Record<string, [Event]>) => {
		const result:Array<EventWithProvider> = []
		for (const provider in events_by_provider) {
			for (const event of events_by_provider[provider]) {
				event.subject = htmlToMarkdown(event.subject)
				result.push({...event, provider: provider})
			}
		}

		result.sort((a,b):number => {
			const at = DateTime.fromISO(a.start.dateTime, {zone: a.start.timeZone})
			const bt = DateTime.fromISO(b.start.dateTime, {zone: b.start.timeZone})

			return at.toMillis() - bt.toMillis()
		})

		return result
	}

	formatEvents = (events:Record<string,[Event]>):string => {
		let result = ""

		const flat_events = this.flattenAndSortEventsByStartTime(events)

		for (const e of flat_events) {
			result += Eta.render(this.plugin.settings.eventTemplate, {...e,  luxon: luxon}) + "\n\n"
		}

		return result
	}

    formatEventsForToday = async () => {
        return this.formatEvents(await this.getEventsForToday())
    }
}