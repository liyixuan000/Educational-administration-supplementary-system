<!--#include file="include/conn.asp"-->
<!--#include file="include/function2.asp"-->
<!--#include file="user_check.asp"-->
<%
Dim class_id
class_id=checkint(trim(request("class_id")),0)
if class_id=0 then
	response.write("Error0")'非法操作
else
	response.write("0,所有形式|")
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from Article002 where pclass_id="&class_id&" "
	rs.open sql,conn,1,2
	do while not rs.eof 
		response.write(rs("id")&","&rs("title_gb")&"|")
		rs.movenext
	loop
	rs.close
	set rs=nothing
end if
%>