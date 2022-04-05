export const defaultMailTemplate =
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