import MSGraphPlugin from "MSGraphPlugin";
import { htmlToMarkdown } from "obsidian";
import { EventWithProvider } from "types";

import { Event } from '@microsoft/microsoft-graph-types'

import { DateTime } from "luxon";

// @ts-ignore
import * as Eta from './node_modules/eta/dist/browser.module.mjs'
import { AppointmentSchema, BasePropertySet, CalendarFolder, CalendarView, DateTime as EWSDateTime, EmailMessageSchema, EwsLogging, ExchangeService, FolderId, ItemView, Mailbox, OAuthCredentials, PropertySet, Uri, WellKnownFolderName } from "ews-javascript-api"

export class CalendarHandler {
    plugin:MSGraphPlugin
	eta:Eta.Eta

    constructor(plugin:MSGraphPlugin) {
        this.plugin = plugin
		this.eta = new Eta.Eta()
    }

    getEventsForTimeRange = async (start:DateTime, end:DateTime):Promise<Record<string, [Event]>> => {
		const events_by_provider:Record<string, [Event]> = {}

		// todo: fetch in parallel using Promise.all()?
		for (const authProviderName in this.plugin.msalProviders) {
			const authProvider = this.plugin.msalProviders[authProviderName]

			if (authProvider.account.type == "MSGraph") {
				const graphClient = this.plugin.getGraphClient(authProvider)

				const dateQuery = `startDateTime=${start.toISO()}&endDateTime=${end.toISO()}`;
				
				const events = await graphClient
					.api('/me/calendarView').query(dateQuery)
					//.select('subject,start,end')
					.orderby(`start/DateTime`)
					.get();

				events_by_provider[authProviderName] = events.value
			} else if (authProvider.account.type == "EWS") {
				const authProvider = this.plugin.msalProviders[authProviderName]

				const token = await authProvider.getAccessToken()

				const svc = new ExchangeService();
				//EwsLogging.DebugLogEnabled = true; // false to turnoff debugging. 
				svc.Url = new Uri(`${authProvider.account.baseUri}/EWS/Exchange.asmx`);
				
				svc.Credentials = new OAuthCredentials(token);

				// Initialize values for the start and end times, and the number of appointments to retrieve.
				
				// Initialize the calendar folder object with only the folder ID. 
				const calendar = await CalendarFolder.Bind(svc, WellKnownFolderName.Calendar, new PropertySet());

				// Set the start and end time and number of appointments to retrieve.
				const calendarView = new CalendarView(new EWSDateTime(start.toMillis()), new EWSDateTime(end.toMillis()));
				
				// Limit the properties returned to the appointment's subject, start time, and end time.
				calendarView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.IsAllDayEvent);
				// Retrieve a collection of appointments by using the calendar view.
				const appointments = await calendar.FindAppointments(calendarView);

				const results = appointments.Items.map(({ Subject, Start, End, IsAllDayEvent }) => (
					{
						subject: Subject, 
						start: {dateTime: Start.ToISOString()}, 
						end: {dateTime: End.ToISOString()}, 
						isAllDay: IsAllDayEvent 
					} as Event)) as [Event];
				
				events_by_provider[authProviderName] = results
			}
		}

		return events_by_provider
	}

	getEventsForToday = async () => {
		const today = DateTime.now().startOf('day')
		const tomorrow = DateTime.now().endOf('day')
		
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

		try {
			for (const e of flat_events) {
				console.log(e)
				result += this.eta.renderString(this.plugin.settings.eventTemplate, {...e,  luxon:DateTime }) + "\n\n"
			}
		} catch (e) {
			console.error(e)
		}

		return result
	}

    formatEventsForToday = async () => {
        return this.formatEvents(await this.getEventsForToday())
    }
}