// Default templates shipped with versions <= 0.9.9. They inserted mail and event content unescaped.
// Kept verbatim (do not edit!) so that unmodified copies stored in data.json can be upgraded to the
// escaping templates automatically.

export const legacyDefaultEventTemplate =
`<%
    formatTime = (dt) => {
	    return moment.utc(dt).local().locale(navigator.language).format("LT")
	}
%>

<% if (it.isAllDay) { %>
### All day: <%~ it.subject %>
<% } else { %>
### <%= formatTime(it.start.dateTime) %>-<%= formatTime(it.end.dateTime) %>: <%~ it.subject %>
<% } %>`

export const legacyDefaultFlaggedMailTemplate =
`<%
    formatMailAddress = (ma) => {
        return ma.name + " <" + ma.address + ">"
    }

    formatDateTime = (dt) => {
        return moment.utc(dt).local().locale(navigator.language).toLocaleString([])
    }

    formatBody = (body) => {
        return body
            .split(/\\r?\\n/)
            .map(s => "\\t" + s)
            .join("\\n")
    }
%>
<% if ('flag' in it) { %>
- [<%~ (it.flag.flagStatus === 'complete') ? 'x' : ' ' %>] <%~ (it.flag.dueDateTime) ? '📅' + formatDateTime(it.flag.dueDateTime.dateTime) : '' %> <%~ it.subject %> [src](<%~ it.webLink %>)
<%~ formatBody(it.bodyPreview) %>
<% } else { %>
- [ ] <%~ it.subject %> [src](<%~ it.webLink %>)
<%~ formatBody(it.bodyPreview) %>
<% } %>
`

export const legacyDefaultMailTemplate =
`<%
    formatMailAddress = (ma) => {
        return ma.name + " <" + ma.address + ">"
    }

    formatDateTime = (dt) => {
        return moment.utc(dt).local().locale(navigator.language).toLocaleString([])
    }

    formatBody = (body) => {
        return body
            .split(/\\r?\\n/)
            .map(s => "> " + s)
            .join("\\n")
    }
%>
> [!NOTE]- <%~ it.subject %> 
> #### From: <%~ formatMailAddress(it.from.emailAddress) %> 
> #### Date: <%~ formatDateTime(it.receivedDateTime) %> 
> #### Body: <%~ formatBody(it.bodyPreview) %> 
`

