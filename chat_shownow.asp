<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.ContentType = "text/html;charset=gb2312"

content = request("content")
application.Lock()
application("Message") = application("Message") & "|" & content
application("u_bound") = application("u_bound") + 1
application.UnLock()
%>
