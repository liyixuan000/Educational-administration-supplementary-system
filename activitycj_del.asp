<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim i,rs,sql
for i = 1 to request("id").count
   if request("id")(i)<>"" and  isnumeric(request("id")(i)) then
   	  set rs=conn.execute("select top 1 * from apply where id="&int(request("id")(i))&"")
	  if not rs.eof then
		'DelFile("/maiyueqi/"&rs("FilePath")&"/"&rs("FileName"))
	  end if
	  rs.close
	  set rs=nothing
      sql = "delete from apply where id="&int(request("id")(i))&""
	  conn.execute (sql)
   end if
next
conn.close
set conn = nothing
response.write "�����ɹ����ѽ�ѡ������Ѷɾ��"
response.write "<meta http-equiv='refresh' content='3;url="&request.servervariables("http_referer")&"'>"
%>