export const defaultEventTemplate =
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