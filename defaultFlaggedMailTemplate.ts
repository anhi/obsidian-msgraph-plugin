// Use <%= ... %> (or it.escapeMarkdown(...)) for anything coming from a mail: it escapes Markdown so
// that a malicious mail cannot inject links, embeds, HTML or code blocks (e.g. dataviewjs) into your
// notes. <%~ ... %> inserts text unchanged and must only be used for already escaped values.
export const defaultFlaggedMailTemplate =
`<%
    const formatDateTime = (dt) => moment.utc(dt).local().locale(navigator.language).toLocaleString([])

    const formatBody = (body) => it.escapeMarkdown(body)
        .split(/\\r?\\n/)
        .map(s => "\\t" + s)
        .join("\\n")
%>
<% if (it.flag) { %>
- [<%~ (it.flag.flagStatus === 'complete') ? 'x' : ' ' %>] <%~ (it.flag.dueDateTime) ? '📅' + it.escapeMarkdown(formatDateTime(it.flag.dueDateTime.dateTime)) : '' %> <%= it.subject %> [src](<%~ it.safeUrl(it.webLink) %>)
<%~ formatBody(it.bodyPreview) %>
<% } else { %>
- [ ] <%= it.subject %> [src](<%~ it.safeUrl(it.webLink) %>)
<%~ formatBody(it.bodyPreview) %>
<% } %>
`
