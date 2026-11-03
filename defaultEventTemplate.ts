// Use <%= ... %> for anything coming from the calendar: it escapes Markdown so that event data
// cannot inject links, embeds, HTML or code blocks into your notes. <%~ ... %> inserts raw text.
export const defaultEventTemplate =
`<%
    const formatTime = (dt) => moment.utc(dt).local().locale(navigator.language).format("LT")
%>

<% if (it.isAllDay) { %>
### All day: <%= it.subject %>
<% } else { %>
### <%= formatTime(it.start.dateTime) %>-<%= formatTime(it.end.dateTime) %>: <%= it.subject %>
<% } %>`
