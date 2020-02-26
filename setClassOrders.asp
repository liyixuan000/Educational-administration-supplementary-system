<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim action,num,class_id,sql
action=request("action")
num=cint(request("num"))
class_id=int(request("class_id"))
if isnumeric(request.form("class_order")(num)) then
	sql="update Articleclass set class_order="&request("class_order")(num)&" where class_id="&class_id&""
	'response.Write(sql)
	'response.End()
	conn.execute (sql)
end if
response.write "<meta http-equiv='refresh' content='0;url="&request.servervariables("http_referer")&"'>"
%>