<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
response.ContentType = "text/html;charset=gb2312"
A_listName=split(Application("UserName"),",")
A_head=split(Application("head"),",")


For i=1 To UBound(A_ListName)%>
	<a href="chat_bottom.asp?name=<%=A_ListName(i)%>" target="bottomFrame">
	<%HeadPath="chat_headICO/"+A_head(i)%>
	<img src="<%=HeadPath%>" border="0" width="23" height="23">
	<%=A_ListName(i)%></a>
	<BR>
<%Next%>
<br>

</p>
