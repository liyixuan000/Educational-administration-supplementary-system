<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.ContentType = "text/html;charset=gb2312"
application.Lock()
u_bound = application("u_bound")
m_bound = session("m_bound")
if u_bound > m_bound then
	A_message=split(Application("Message"),"|")
	tmpChat = ""

	for i=m_bound+1 to u_bound
	tmpChat = tmpChat + A_message(i) + "<BR>"
	Next
	session("m_bound") = u_bound
end if
application.UnLock()
response.Write tmpChat
%>