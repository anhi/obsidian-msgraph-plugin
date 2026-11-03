// Use <%= ... %> (or it.escapeMarkdown(...)) for anything coming from a mail: it escapes Markdown so
// that a malicious mail cannot inject links, embeds, HTML or code blocks (e.g. dataviewjs) into your
// notes. <%~ ... %> inserts text unchanged and must only be used for already escaped values.
export const defaultMailTemplate =
`<%
    const formatMailAddress = (ma) => ma ? it.escapeMarkdown((ma.name || "") + " <" + (ma.address || "") + ">") : ""

    const formatDateTime = (dt) => moment.utc(dt).local().locale(navigator.language).toLocaleString([])

    const formatBody = (body) => it.escapeMarkdown(body)
        .split(/\\r?\\n/)
        .map(s => "> " + s)
        .join("\\n")
%>
> [!NOTE]- <%= it.subject %> 
> #### From: <%~ formatMailAddress(it.from?.emailAddress) %> 
> #### Date: <%= formatDateTime(it.receivedDateTime) %> 
> #### Body: <%~ formatBody(it.bodyPreview) %> 
`
