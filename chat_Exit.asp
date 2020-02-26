<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
A_UserName=Split(Application("UserName"),",")
A_head=Split(Application("head"),",")
UserName=Session("UserName")
Application.Lock()
Application("UserName")=" "
Application("head")=" "
Application("UserName")=A_UserName(0)
Application("head")=A_head(0)
For i=1 To UBound(A_UserName)
	If A_UserName(i)<> UserName Then
		Application("UserName")=Application("UserName")+","+A_UserName(i)
		Application("head")=Application("head")+","+A_head(i)
	End If
Next
Application.UnLock()
Session.Abandon()
%>
<script language="javascript">
alert("愿你每天开心！");
document.location.href="chat_index.asp"
</script>
