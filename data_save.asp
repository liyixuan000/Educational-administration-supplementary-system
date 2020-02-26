<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
Dim action:action=request("action")
Dim Class_id,title,pclass_id
Class_id=checkint(request("Class_id"),0)
pclass_id=checkint(request("pclass_id"),0)
if Class_id=0 then jstop("温馨提示：请选择类别！")
isTuiJian =checkint(request("isTuiJian"),0)
orders =checkint(request("orders"),0)
Keyword=Request("keyword")
Postdate=now()
picUrl=saferequest("picUrl")
if action="Add" then
	call Add()
elseIf action="Edit" then
	Call Edit()
end if
Sub Add()
	dim rs,sqlstr
	sqlstr="select * from Article002"
	set rs =server.CreateObject("adodb.recordset")
	rs.cursorlocation = 3 
	rs.open sqlstr,conn,1,3
	rs.addnew
	  rs("title_gb")=trim(Request.Form("title_gb"))
	  rs("content_gb")=trim(Request.Form("content_gb"))
	  rs("pagetitle_gb")=trim(Request.Form("pagetitle_gb"))
	rs("Class_id")=Class_id
	rs("pclass_id")=pclass_id
	rs("picUrl")=picUrl
	rs("addtimes")=Postdate
	rs("username")=session("loginusername")
	'rs("language")=language
	rs.update
	'dim News_ID
	'News_ID = rs("id")
	rs.close
	set rs = nothing
	Call JsAlertAndGoto_1("温馨提示：添加成功！需要继续添加吗？","data_add.asp?lag="&language,"data_manage.asp?lag="&language)
End Sub

Sub Edit()
	dim rs,sqlstr,Id
	Id =checkint(request("id"),0)
	if id=0 then jstop("温馨提示：参数错误！")
	sqlstr="select * from Article002 where id="&Id
	set rs =server.CreateObject("adodb.recordset")
	rs.open sqlstr,conn,1,3
	  rs("title_gb")=trim(Request.Form("title_gb"))
	  rs("content_gb")=trim(Request.Form("content_gb"))
	  rs("pagetitle_gb")=trim(Request.Form("pagetitle_gb"))
	rs("Class_id")=Class_id
	rs("pclass_id")=pclass_id
	rs("picUrl")=picUrl
	rs("addtimes")=Postdate
	rs("username")=session("loginusername")
	rs.update
	rs.close
	set rs = nothing
	Call JsAlertAndGoto("温馨提示：修改成功！","data_manage.asp?&page="&request("page")&"&lag="&request("lag"))
End Sub


%>