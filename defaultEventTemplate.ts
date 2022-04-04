export const defaultEventTemplate =
`<%
    formatTime = (ds) => {
	    return new Date(ds).toLocaleTimeString([], {hour: 'numeric', minute: 'numeric'})
	}
%>

### <%= formatTime(it.start.dateTime) %>-<%= formatTime(it.end.dateTime) %>: <%~ it.subject %>`