import { Eta } from "eta/core"

/**
 * Characters that can change the meaning of Markdown in Obsidian:
 *  - code fences / inline code (`, ~)  -> would allow injecting ```dataviewjs blocks or `$= ...` inline JS
 *  - HTML (<, >)                       -> iframes, Templater tags (<% ... %>), block quotes / callouts
 *  - links & embeds ([, ], !)          -> wikilinks, remote image embeds (tracking pixels)
 *  - emphasis, headings, tables, tags, math, comments, highlights, block ids (* _ # | $ % = ^)
 *  - the escape character itself (\)
 */
const MARKDOWN_SPECIAL = /[\\`*_~[\]<>#!|$%=^{}]/g

/**
 * Escapes a value coming from an untrusted source (mail subject, mail body, event subject, …)
 * so that it is rendered literally by Obsidian. Line breaks are preserved.
 */
export const escapeMarkdown = (value: unknown): string => {
    if (value === undefined || value === null) return ""
    return String(value).replace(MARKDOWN_SPECIAL, (c) => "\\" + c)
}

/**
 * Returns the URL if it is an http(s) URL, percent-encoding characters that would terminate a
 * Markdown link destination. Returns an empty string for anything else (javascript:, file:, obsidian:, …).
 */
export const safeUrl = (value: unknown): string => {
    if (typeof value !== "string") return ""
    let url: URL
    try {
        url = new URL(value)
    } catch {
        return ""
    }
    if (url.protocol !== "https:" && url.protocol !== "http:") return ""
    return url.href.replace(/[\s()<>[\]`]/g, (c) => "%" + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0"))
}

/**
 * Shared template engine.
 *  - `<%= value %>` escapes Markdown (use it for anything that comes from a mail or an event)
 *  - `<%~ value %>` inserts the value unchanged
 */
export const eta = new Eta({
    autoEscape: true,
    escapeFunction: escapeMarkdown,
})

/** Helpers that are available to every template as `it.<helper>`. */
export const templateHelpers = {
    escapeMarkdown,
    safeUrl,
}
