import { Client, GraphRequest } from "@microsoft/microsoft-graph-client"
import { DateTime } from "luxon"

export const formatTime = (date:Date): string => {
    const dt = DateTime.fromJSDate(date)
    return dt.toLocaleString(DateTime.TIME_SIMPLE)
}

export const errorMessage = (e: unknown): string => {
    if (e instanceof Error) return e.message
    if (typeof e === 'object' && e !== null && 'message' in e) return String((e as {message: unknown}).message)
    return String(e)
}

/**
 * Executes a Graph collection request and follows @odata.nextLink until `maxItems` items
 * have been collected or there are no more pages.
 */
export const getAllPages = async <T>(client: Client, request: GraphRequest, maxItems: number): Promise<T[]> => {
    let response = await request.get()
    const items: T[] = [...(response?.value ?? [])]

    while (response?.['@odata.nextLink'] && items.length < maxItems) {
        response = await client.api(response['@odata.nextLink']).get()
        items.push(...(response?.value ?? []))
    }

    return items.slice(0, maxItems)
}
