<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
Dim action:action=request("action")
Dim pid,title,Postdate,bnumber,prize
pid=checkint(request("pid"),0)
'if pid=0 then jstop("温馨提示：请选择活动！")
orders =checkint(request("orders"),0)
Postdate=saferequest("Postdate")
bnumber=saferequest("bnumber")
prize=saferequest("prize")
picUrl=saferequest("picUrl")
if action="Add" then
	call Add()
elseIf action="Edit" then
	Call Edit()
end if
Sub Add()
	dim rs,sqlstr
	sqlstr="select * from photo"
	set rs =server.CreateObject("adodb.recordset")
	rs.cursorlocation = 3 
	rs.open sqlstr,conn,1,3
	rs.addnew
	  rs("bnumber")=trim(Request.Form("bnumber"))
	  rs("prize")=trim(Request.Form("prize"))
	rs("pid")=pid
	rs("picUrl")=picUrl
	rs("orders")=orders
	rs("username")=session("loginusername")
	rs("addtimes")=now()
	rs.update
	'dim News_ID
	'News_ID = rs("id")
	rs.close
	set rs = nothing
	Call JsAlertAndGoto_1("温馨提示：添加成功！需要继续添加吗？","photo_Add.asp?lag="&language,"photo_Manage.asp?lag="&language)
End Sub

Sub Edit()
	dim rs,sqlstr,Id
	Id =checkint(request("id"),0)
	if id=0 then jstop("温馨提示：参数错误！")
	sqlstr="select * from photo where id="&Id
	set rs =server.CreateObject("adodb.recordset")
	rs.open sqlstr,conn,1,3
	  rs("bnumber")=trim(Request.Form("bnumber"))
	  rs("prize")=trim(Request.Form("prize"))
	rs("pid")=pid
	rs("picUrl")=picUrl
	rs("orders")=orders
	'rs("username")=session("loginusername")
	rs("addtimes")=now()
	rs.update
	rs.close
	set rs = nothing
	Call JsAlertAndGoto("温馨提示：修改成功！","photo_Manage.asp?&page="&request("page")&"&lag="&request("lag"))
End Sub


%>