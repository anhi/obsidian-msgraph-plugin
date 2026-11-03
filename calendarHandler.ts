import { Notice } from "obsidian"
import { Event } from '@microsoft/microsoft-graph-types'
import { DateTime } from "luxon"
import {
	AppointmentSchema,
	CalendarFolder,
	CalendarView,
	DateTime as EWSDateTime,
	ExchangeService,
	OAuthCredentials,
	PropertySet,
	Uri,
	WellKnownFolderName,
} from "ews-javascript-api"

import type MSGraphPlugin from "./MSGraphPlugin"
import type { MSALAuthProvider } from "./authProvider"
import { normalizeEwsBaseUri } from "./authConfig"
import { eta, templateHelpers } from "./templating"
import { EventWithProvider } from "./types"
import { errorMessage, getAllPages } from "./util"

/** Upper bound for the number of events fetched per account and time range. */
const MAX_EVENTS = 1000

export class CalendarHandler {
	plugin: MSGraphPlugin

	constructor(plugin: MSGraphPlugin) {
		this.plugin = plugin
	}

	private getGraphEvents = async (authProvider: MSALAuthProvider, start: DateTime, end: DateTime): Promise<Event[]> => {
		const graphClient = authProvider.graphClient
		const dateQuery = `startDateTime=${encodeURIComponent(start.toISO() ?? "")}&endDateTime=${encodeURIComponent(end.toISO() ?? "")}`

		const request = graphClient
			.api('/me/calendarView').query(dateQuery)
			.orderby(`start/DateTime`)
			.top(100)

		return getAllPages<Event>(graphClient, request, MAX_EVENTS)
	}

	private getEwsEvents = async (authProvider: MSALAuthProvider, start: DateTime, end: DateTime): Promise<Event[]> => {
		const baseUri = normalizeEwsBaseUri(authProvider.account.baseUri)
		const token = await authProvider.getAccessToken()

		const svc = new ExchangeService()
		svc.Url = new Uri(`${baseUri}/EWS/Exchange.asmx`)
		svc.Credentials = new OAuthCredentials(token)

		// Initialize the calendar folder object with only the folder ID.
		const calendar = await CalendarFolder.Bind(svc, WellKnownFolderName.Calendar, new PropertySet())

		const calendarView = new CalendarView(new EWSDateTime(start.toMillis()), new EWSDateTime(end.toMillis()), MAX_EVENTS)

		// Limit the properties returned to the appointment's subject, start time, and end time.
		calendarView.PropertySet = new PropertySet(AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.IsAllDayEvent)
		const appointments = await calendar.FindAppointments(calendarView)

		return appointments.Items.map(({ Subject, Start, End, IsAllDayEvent }) => ({
			subject: Subject,
			start: {dateTime: Start.ToISOString()},
			end: {dateTime: End.ToISOString()},
			isAllDay: IsAllDayEvent,
		} as Event))
	}

	/**
	 * Fetches the events of all enabled accounts. Accounts that fail are reported and skipped.
	 */
	getEventsForTimeRange = async (start: DateTime, end: DateTime): Promise<Record<string, Event[]>> => {
		const eventsByProvider: Record<string, Event[]> = {}

		for (const [name, authProvider] of this.plugin.msalProviders) {
			if (!authProvider.account.enabled) continue
			try {
				eventsByProvider[name] = authProvider.account.type === "EWS"
					? await this.getEwsEvents(authProvider, start, end)
					: await this.getGraphEvents(authProvider, start, end)
			} catch (e) {
				new Notice(`MSGraph: could not fetch events for "${name}": ${errorMessage(e)}`)
				console.error(e)
			}
		}

		return eventsByProvider
	}

	getEventsForToday = async () => {
		return await this.getEventsForTimeRange(DateTime.now().startOf('day'), DateTime.now().endOf('day'))
	}

	flattenAndSortEventsByStartTime = (eventsByProvider: Record<string, Event[]>): EventWithProvider[] => {
		const result: EventWithProvider[] = []
		for (const provider in eventsByProvider) {
			for (const event of eventsByProvider[provider]) {
				result.push({...event, provider})
			}
		}

		const startMillis = (e: Event): number => {
			const dateTime = e.start?.dateTime ?? ""
			let dt = DateTime.fromISO(dateTime, {zone: e.start?.timeZone ?? undefined})
			if (!dt.isValid) dt = DateTime.fromISO(dateTime, {zone: "utc"})
			return dt.isValid ? dt.toMillis() : 0
		}

		return result.sort((a, b) => startMillis(a) - startMillis(b))
	}

	formatEvents = (events: Record<string, Event[]>): string => {
		let result = ""

		for (const e of this.flattenAndSortEventsByStartTime(events)) {
			try {
				result += eta.renderString(this.plugin.settings.eventTemplate, {...e, ...templateHelpers, luxon: DateTime}) + "\n\n"
			} catch (err) {
				new Notice(`MSGraph: error in event template: ${errorMessage(err)}`)
				console.error(err)
				break
			}
		}

		return result
	}

	formatEventsForToday = async () => {
		return this.formatEvents(await this.getEventsForToday())
	}
}
