import { DateTime } from "luxon"

export const formatTime = (date:Date): string => {
    const dt = DateTime.fromJSDate(date)
    return dt.toLocaleString(DateTime.TIME_SIMPLE)
}