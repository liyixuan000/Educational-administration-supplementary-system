<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
Dim action:action=request("action")
dim names,phone,email,content,power,username,password
power =checkint(request("power"),0)
names=saferequest("name")
phone=saferequest("phone")
email=saferequest("email")
content=saferequest("content")
username=saferequest("username")
password=saferequest("password")
if action="add" then
	call Add()
elseIf action="edit" then
	Call Edit()
end if
Sub Add()
	dim rs,sqlstr
	sqlstr="select * from Userinfo where username='"&username&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sqlstr,conn,1,3
	if not rs.eof then
		jstop("�û����Ѿ�����!")
	else
		rs.addnew
		rs("username")=username
		rs("password")=password
		rs("name")=names
		rs("phone")=phone
		rs("email")=email
		rs("content")=content
		rs("power")=power
		rs.update
	end if
	rs.close
	set rs=nothing
	Call JsAlertAndGoto_1("��ܰ��ʾ����ӳɹ�����Ҫ���������","user_add.asp","user_manage.asp?page="&request("page"))
End Sub

Sub Edit()
	dim rs,sqlstr,Id
	Id =checkint(request("id"),0)
	if id=0 then jstop("��ܰ��ʾ����������")
	sqlstr="select * from userinfo where id="&Id
	set rs =server.CreateObject("adodb.recordset")
	rs.open sqlstr,conn,1,3
	if password<>"" then rs("password")=password
	rs("name")=names
	rs("phone")=phone
	rs("email")=email
	rs("content")=content
	rs("power")=power
	rs.update
	rs.close
	set rs = nothing
	Call JsAlertAndGoto("��ܰ��ʾ���޸ĳɹ���","user_my_edit.asp")
End Sub

%>