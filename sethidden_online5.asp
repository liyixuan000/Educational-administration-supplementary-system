<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim i,sql
for i = 1 to request("id").count
   if request("id")(i)<>"" and  isnumeric(request("id")(i)) then
   	  if request("action")="hidden" then
      sql = "update honor set isok=0 where id="&int(request("id")(i))&""
	  elseif request("action")="online" then
	  	sql = "update honor set isok=1 where id="&int(request("id")(i))&""
	  end if
	  conn.execute (sql)
   end if
next
conn.close
set conn = nothing
response.write "操作成功，已将选定的记录重新设置"
response.write "<meta http-equiv='refresh' content='3;url="&request.servervariables("http_referer")&"'>"
%>