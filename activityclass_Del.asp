<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
sub delClass(class_id)
	dim sqlc0,rsc0
	sqlc0="select class_id,childSum from Articleclass where parent_id="&class_id&" "
	set rsc0=server.createobject("adodb.recordset")
	rsc0.open sqlc0,conn,1,3
	do while not rsc0.eof
		class_id=rsc0("class_id")
		conn.execute ("delete from Article where class_id="&class_id&" ")
		call delClass(class_id)
		rsc0.delete
		rsc0.update
		rsc0.movenext
	loop
	rsc0.close:set rsc0=nothing
end sub
dim class_id:class_id=checkint(request("class_id"),0)
if class_id=0 then
	response.end
end if

dim sql,rs,parent_id
sql="select class_id,parent_id from Articleclass where class_id="&class_id&""
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
if not rs.eof then
	parent_id=rs("parent_id")
	conn.execute ("update Articleclass set childSum=childSum-1 where class_id="&parent_id&"")
	conn.execute ("delete from Article where class_id="&class_id&"")
	call delClass(class_id)
	rs.delete
	rs.update
end if
rs.close:set rs=nothing
conn.close:set conn = nothing
response.write "操作成功，已将选定的记录删除"
response.write "<meta http-equiv='refresh' content='3;url="&request.servervariables("http_referer")&"'>"
%>