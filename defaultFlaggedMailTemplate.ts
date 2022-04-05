export const defaultFlaggedMailTemplate =
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