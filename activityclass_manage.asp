<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
if session("power")<>2 then
	response.Write("对不起,您没有权限操作此页面")
	response.End()
end if
sub subClass(class_id)
	dim rsChildSum,counter,sql,countS,strImg,rs,nodeImg,i
	set rsChildSum=server.CreateObject("adodb.recordset")
	rsChildSum.open "select * from Articleclass where parent_ID="&class_ID,conn,1,1
	counter=rsChildSum.recordcount
	rsChildSum.close:set rsChildSum=nothing
	sql="select * from Articleclass where parent_id="&class_id&"   order by class_order desc,class_id asc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	countS=0
	do while not rs.eof
		countS=countS+1
		num=num+1
		strImg=""
		if rs("childSum")>0 then
			nodeImg="<img src='images/treeOpen.gif' align='absmiddle'>"
		else
			nodeImg="<img src='images/treeNode.gif' align='absmiddle'>"
		end if
		for i=1 to rs("class_depth")
			if i>=rs("class_Depth") then 
				if countS>=counter then strImg=strImg&"<img src='images/treeBottom.gif' align='absmiddle'>" else strImg=strImg&"<img src='images/treeMiddle.gif' align='absmiddle'>"
			else
				strImg=strImg&"<img src='images/treeLeft.gif' align='absmiddle'>"
			end if
		next
		response.write "<tr>"

		response.write "<td bgcolor='#ffffff' style='padding-left:6px;' height=20>"&strImg&""&nodeImg&"&nbsp;"&rs("class_name_"&lan_txt&"")&"</td>"
		'response.write "<td bgcolor='#ffffff' align=center><input name='class_order' type='text' id='class_order' value='"&rs("class_order")&"' size='6' maxlength='6'><input type='button' name='Submit' value='更新' onClick=""setClassOrders(this.form,"&rs("class_id")&","&num&",'"&request("lag")&"')""></td>"
		response.write "<td bgcolor='#ffffff' align=center><a href=activityclass_add.asp?class_id="&rs("class_id")&"&lag="&request("lag")&">添加子类</a>&nbsp;&nbsp;<a href=activityclass_Edit.asp?class_id="&rs("class_id")&"&parent_id_Old="&rs("parent_id")&">修改</a>&nbsp;&nbsp;<a href='activityclass_del.asp?class_id="&rs("class_id")&"' onclick='return checkDelClass()'>删除</a></td>"
		response.write "</tr>"
		class_id=rs("class_id")
		call subClass(class_id)
		rs.movenext
	loop
	rs.close:set rs=nothing
end sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
<link rel="stylesheet" href="css/css.css"  type="text/css">
<script language="javascript" src="act_check.js"></script>
</head>
<body>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="35" bgcolor="#ECF5FF">&nbsp;欢迎您，<font color="#FF0000"><%=session("loginusername")%></font>,您现在的身份是<font color="#FF0000"><%if session("power")=1 then
	response.Write("超级管理员")
	elseif session("power")=2 then
	response.Write("普通管理员")
	else
	response.Write("学生")
	end if%></font></td>
  </tr>
</table><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100" height="35" bgcolor="#FFFFFF">&nbsp;<strong>活动分类：<font color="#FF0000"></font></strong></td>
    <td bgcolor="#FFFFFF"><table  border="0" cellspacing="1" cellpadding="1">
      <tr>        <%
				dim rsc888,selected888
			 	set rsc888=conn.execute ("select class_id,class_name_"&lan_txt&",class_depth from Articleclass where parent_id=0  order by class_order desc,class_id asc")
				do while not rsc888.eof

			 %>
        <td width="80" height="35" align="center" <%if rsc888("class_id")=class_id then 
		response.Write("bgcolor='#1496F5'")
		else
		response.Write("bgcolor='#C8EDFB'")
		end if
		%>
		><a href="activity_manage.asp?class_id=<%=rsc888("class_id")%>"><%=rsc888("class_name_"&lan_txt&"")%></a></td>
        <%

					rsc888.movenext
				loop
				rsc888.close:set rsc888=nothing
				%>

      </tr>
    </table></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td width="120" valign="top">        			<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
    <td height="25" bgcolor="#39B2DD"><a href="activity_manage.asp" target="main">
					<span >
					<font color="#fff">活动列表</font></span></a></td>
  </tr><%if session("power")=3 then%>
  <tr>
    <td height="25" bgcolor="#39B2DD">
					<a href="activitycj_add.asp" target="main"><font color="#fff">
					<span >申请活动成绩认证</span></font></a></li>
                   </td>
  </tr> <%end if%>
  <tr>
    <td height="25" bgcolor="#39B2DD"><a href="activitycj_manage.asp" target="main">
					<font color="#fff">
					<span >活动认证申请列表</span></font></a></td>
  </tr><%if session("power")=2 then%>
  <tr>
    <td height="25" bgcolor="#39B2DD"><a href="activity_add.asp" target="main"><font color="#fff">
					<span >添加活动</span></font></a></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#39B2DD"><a href="activityclass_add.asp" target="main"><font color="#fff">
					<span >添加活动分类</span></font></a></td>
  </tr>
  <tr>
    <td height="25" bgcolor="#39B2DD"><a href="activityclass_manage.asp" target="main"><font color="#fff">
					<span >管理活动分类</span></font></a></td>
  </tr><%end if%>
                    </table>

					</td>
    <td width="567" valign="top" bgcolor="#ECF5FF">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td bgcolor="#ECF5FF">
<div align="center"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" align="left"> 活动类别管理 </td>
        <td></td>
      </tr>
  </table>

<form name="form1" method="post" action="">
  <table width="95%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><input type="button" name="Submit" value="添加类别" onClick="window.location.href='activityclass_add.asp?lag=<%=request("lag")%>'"></td>
      <td align="right">&nbsp;</td>
    </tr>
  </table>
  <table cellpadding="0" cellspacing="1" border="0" width="95%" align="center" class="table_southidc">
	<tr>
		<td class="back_southidc" height="25" align="center">活动类别管理</td>
	</tr>
	<tr>
	<td><table width="100%" height="47"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
    <tr align="center" bgcolor="#E6E6E6">
	        
      <td width="35%">活动类别 </td>
      <!--
      <td width="28%">排列顺序</td>-->
      <td width="37%">操作</td>
    </tr>
    <%
	dim rs2,sql2
	  dim rsCounter0,counter0,sql,class_id,count0S,rs0,num,nodeImg0,strImg0
	set rsCounter0=server.CreateObject("adodb.recordset")
	rsCounter0.open "select * from Articleclass where parent_id=0",conn,1,1
	counter0=rsCounter0.recordcount
	rsCounter0.close:set rsCounter0=nothing
	  sql = "select * from Articleclass where parent_id=0  order by class_order desc,class_id asc"
	  set rs0 = conn.execute (sql)
	  num=0               '排列顺序号的变量
	  count0S=0
	  do while not rs0.eof
	  num=num+1
	  count0S=count0S+1
	  nodeImg0=""
	  strImg0=""
	  if rs0("childSum")>0 then nodeImg0="<img src='images/treeOpen.gif' align='absmiddle'>" else nodeImg0="<img src='images/treeNode.gif' align='absmiddle'>"
	  if count0S>=counter0 then strImg0="<img src='images/treeBottom.gif' align='absmiddle'>" else strImg0="<img src='images/treeMiddle.gif' align='absmiddle'>"
	  %>
    <tr bgcolor="#FFFFFF">

	        <td height="18" style="padding-left:6px;"><%=strImg0%><%=nodeImg0%>&nbsp;<%=rs0("class_name_"&lan_txt&"")%></td>

     <!-- <td align="center"><input name="class_order" type="text" id="class_order" value="<%=rs0("class_order")%>" size="6" maxlength="6">
        <input type="button" name="Submit" value="更新" onClick="setClassOrders(this.form,<%=rs0("class_id")%>,<%=num%>,'<%=request("lag")%>')"></td>-->
      <td align="center" bgcolor="#FFFFFF"><a href=activityclass_add.asp?class_id=<%=rs0("class_id")%>&lag=<%=request("lag")%>>添加子类</a>&nbsp;&nbsp;<a href="Activityclass_Edit.asp?class_id=<%=rs0("class_id")%>&parent_id_Old=<%=rs0("parent_id")%>&lag=<%=request("lag")%>">修改</a>&nbsp;&nbsp;<a href="activityclass_del.asp?class_id=<%=rs0("class_id")%>&lag=<%=request("lag")%>" onClick="return checkDelClass()">删除</a></td>
    </tr>
    <%
	  	class_id=rs0("class_id")
	  	call subClass(class_id)
	  	rs0.movenext
		loop
		if rs0.eof and rs0.bof then response.write "<tr><td height=32 colspan=4 bgcolor=#ffffff>没有类别</td></tr>"
		rs0.close:set rs0 = nothing
		
	  %>
  </table>
  </td>
  </tr>
  </table>
  </form></div>
</td>

</tr>
<tr>
<td></td>
</tr>
</table>
	</td>
  </tr>
  <tr>
    <td width="778" colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</div>
</body>
</html>
<script>
function check()
{
	if(myform.username.value=='')
	{
		alert("用户名不能为空！");
		myform.username.focus();
		return false;
	}
	if(myform.password.value=='')
	{
		alert("密码不能为空！");
		myform.password.focus();
		return false;
	}
	if(myform.repassword.value=='')
	{
		alert("确认密码不能为空！");
		myform.repassword.focus();
		return false;
	}
	if(myform.repassword.value!=myform.password.value)
	{
		alert("两次输入的密码不一致！");
		myform.repassword.focus();
		return false;
	}
	if(myform.name.value=='')
	{
		alert("姓名不能为空！");
		myform.name.focus();
		return false;
	}
	if(myform.phone.value=='')
	{
		alert("联系方式不能为空！");
		myform.phone.focus();
		return false;
	}
}
</script>