<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
if session("power")<>2 then
	response.Write("对不起,您没有权限操作此页面")
	response.End()
end if
Sub AddClass()
	dim parent_id,class_order
	parent_id=checkint(request("parent_id"),0)
	class_order =checkint(request("class_order"),0)
	dim sqlstr,rs,class_depth
	set rs = server.createobject("adodb.recordset")
	sqlstr="select * from Articleclass "
	rs.open sqlstr,conn,1,3
	rs.addnew	
	 rs("class_name_gb")=trim(Request.Form("class_name_gb"))
	rs("parent_id")=parent_id
	rs("class_order")=class_order
	if parent_id=0 then
		 rs("class_depth")=1
	else
		dim sql0,rs0
		sql0="select class_depth,childSum from Articleclass where class_id="&parent_id&""
		set rs0=server.createobject("adodb.recordset")
		rs0.open sql0,conn,1,3
		rs0("childSum")=rs0("childSum")+1
		rs0.update
		class_depth=rs0("class_depth")+1
		rs0.close:set rs0=nothing
		rs("class_depth")=class_depth
	end if
	rs.update
	rs.close
	set rs=nothing
	response.write "<script>alert('添加成功');window.location.href='activityclass_Manage.asp?lag="&request("lag")&"'</script>"
	response.end
End Sub
sub subClass(class_id)
	dim rscounter,counter,rs,countS,selected,i,nbsp
	set rsCounter=server.CreateObject("adodb.recordset")
	rsCounter.open "select * from Articleclass where parent_id="&class_id&" ",conn,1,1
	counter=rsCounter.recordcount
	rsCounter.close:set rsCounter=nothing
	sql="select parent_id,class_id,class_name_"&lan_txt&",class_order,class_depth,childSum from Articleclass where parent_id="&class_id&"  order by class_order desc,class_id asc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	countS=0
	do while not rs.eof
		countS=countS+1
		nbsp=""
		if request("class_id")=cstr(rs("class_id")) then
			selected="selected"
		else
			selected=""
		end if
		for i=1 to rs("class_depth")
			if i>=rs("class_depth") then
				if countS>=counter then nbsp=nbsp&"┗&nbsp;&nbsp;" else nbsp=nbsp&"┝&nbsp;&nbsp;"
			else
				nbsp=nbsp&"┃&nbsp;&nbsp;"
			end if
		next
		response.write "<option value='"&rs("class_id")&"' "&selected&">"&nbsp&""&rs("class_name_"&lan_txt&"")&"</option>"
		class_id=rs("class_id")
		call subClass(class_id)
		rs.movenext
	loop
	rs.close:set rs=nothing
end sub
if request.form("action")="add" then
	Call AddClass()
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
<link rel="stylesheet" href="css/css.css"  type="text/css">
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
    <td width="567" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td bgcolor="#99CCFF">
<div align="center"><form name="form1" method="post" action="" onSubmit="javascript:{if(this.class_name.value.replace(/\s/gi,'')==''){alert('请输入类别名称');this.class_name.focus();return false;}}">
  <table cellpadding="0" cellspacing="1" border="0" width="95%" align="center" class="table_southidc">
	<tr>
		<td class="back_southidc" height="25" align="center">活动类别添加</td>
	</tr>
	<tr>
	<td>
  <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1" >
    		  <tr><td colspan="2">

		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="tbbg01">
            <td height="25" align="right" valign="top">类别名称：</td>
            <td align="left" valign="top" class="tbbg01"><input name="class_name_gb" type="text" id="class_name_gb" size="50" maxlength="100">            </td>
          </tr>	         

</table>				
</td></tr>
    <tr bgcolor="#FFFFFF">
      <td height="37" align="right"  class="tbbg01" width="150" >所属分类：</td>
      <td bgcolor="#FFFFFF"  class="tbbg01" width="850">&nbsp;
		 <select name="parent_id" id="parent_id">
		<%
		dim rs0,sql,selected,class_id
		response.write "<option value='0'>（无）作为一级分类</option>"
		sql="select class_id,parent_id,class_name_"&lan_txt&",class_order from Articleclass where parent_id=0  order by class_order desc,class_id asc"
		set rs0=server.createobject("adodb.recordset")
		rs0.open sql,conn,1,1
		do while not rs0.eof
			if request("class_id")=cstr(rs0("class_id")) then
			selected="selected"
		else
			selected=""
		end if
			response.write "<option value='"&rs0("class_id")&"' "&selected&">"&rs0("class_name_"&lan_txt&"")&"</option>"
			class_id=rs0("class_id")
			call subClass(class_id)
		rs0.movenext
		loop
		rs0.close:set rs0=nothing
		%>
      </select></td>
    </tr><!--
    <tr  class="tbbg01">
      <td height="34" align="right" >排序顺序：</td>
      <td >&nbsp;
        <input name="class_order" type="text" id="class_order" size="6" maxlength="6">
      <span class="style1">（该值为一个数值，数值越大类别越排前面）        </span></td>
    </tr>-->
    <tr  class="tbbg01">
      <td height="36"  class="tbbg01">&nbsp;</td>
      <td height="36" align="left"  class="tbbg01" >&nbsp;
        <input type="submit" name="Submit" value=" 添 加 " /><input type="hidden" name="action" value="add" /></td></tr>
  </table>
 </td>
 </tr>
 </table>
</form></div>
</td>
<td></td>
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
