<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
if session("power")<>2 then
	response.Write("对不起,您没有权限操作此页面")
	response.End()
end if
dim class_id,parent_id_Old
class_id=checkint(request("class_id"),0)
parent_id_Old=checkint(request("parent_id_Old"),0)
if class_id=0 then 
	jstop("温馨提示：参数错误！")
end if

Sub EditClass()
	dim rs,sql,class_order,parent_id,childSum,class_depth
	class_order =checkint(request("class_order"),0)
	parent_id=checkint(request("parent_id"),0)
	sql="select class_depth,childSum from Article002class where class_id="&parent_id&""
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then 
	class_depth=rs("class_depth")+1
	else
	class_depth=1
	end if
	sql="select * from Article002class where class_id="&class_id&""
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs("class_name_gb")=trim(Request.Form("class_name_gb"))
	rs("Class_order")=Class_order
	rs("parent_id")=parent_id
	rs("class_depth")=class_depth
	rs.update
	rs.close
	set rs=nothing
	response.write "<script>alert('修改成功');window.location.href='dataclass_Manage.asp?lag="&request("lag")&"'</script>"
	response.end
End Sub

sub subClass(class_id,parent_id)
	dim rscounter,counter,rs,countS,selected,i,nbsp
	set rsCounter=server.CreateObject("adodb.recordset")
	rsCounter.open "select * from Article002class where parent_id="&class_id&" ",conn,1,1
	counter=rsCounter.recordcount
	rsCounter.close:set rsCounter=nothing
	sql="select parent_id,class_id,class_name_"&lan_txt&",class_order,class_depth,childSum from Article002class where parent_id="&class_id&"  order by class_order desc,class_id asc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	countS=0
	do while not rs.eof
		countS=countS+1
		nbsp=""
		if parent_id=rs("class_id") then
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
		call subClass(class_id,parent_id)
		rs.movenext
	loop
	rs.close:set rs=nothing
end sub
if request("action")="edit" then call EditClass()
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
    <td width="100" height="35" bgcolor="#FFFFFF">&nbsp;<strong>资料分类：<font color="#FF0000"></font></strong></td>
    <td bgcolor="#FFFFFF"><table  border="0" cellspacing="1" cellpadding="1">
      <tr>        <%
				dim rsc888,selected888
			 	set rsc888=conn.execute ("select class_id,class_name_"&lan_txt&",class_depth from Article002class where parent_id=0  order by class_order desc,class_id asc")
				do while not rsc888.eof

			 %>
        <td width="80" height="35" align="center" <%if rsc888("class_id")=class_id then 
		response.Write("bgcolor='#1496F5'")
		else
		response.Write("bgcolor='#C8EDFB'")
		end if
		%>
		><a href="data_manage.asp?class_id=<%=rsc888("class_id")%>"><%=rsc888("class_name_"&lan_txt&"")%></a></td>
        <%

					rsc888.movenext
				loop
				rsc888.close:set rsc888=nothing
				%><td width="80" height="35" align="center" bgcolor='#C8EDFB'><a href="person_manage.asp">个人信息</a></td>

      </tr>
    </table></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td width="120" valign="top">        			<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="data_manage.asp" target="main"> <span > <font color="#fff">资料列表</font></span></a></td>
      </tr>
      <%if session("power")=3 then%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="person_add.asp" target="main"><font color="#fff"> <span >发布个人信息</span></font></a>
          </li></td>
      </tr>
      <%end if%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="person_manage.asp" target="main"> <font color="#fff"> <span >个人信息发布列表</span></font></a></td>
      </tr>      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="data_add.asp" target="main"><font color="#fff"> <span >添加资料</span></font></a></td>
      </tr>
      <%if session("power")=2 then%>

      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="dataclass_add.asp" target="main"><font color="#fff"> <span >添加资料分类</span></font></a></td>
      </tr>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="dataclass_manage.asp" target="main"><font color="#fff"> <span >管理资料分类</span></font></a></td>
      </tr>
      <%end if%>
    </table>

					</td>
    <td width="567" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td bgcolor="#99CCFF"><%dim rsEdit
		set rsEdit = conn.execute ("select * from Article002class where Class_id="&Class_id&"")
		if not rsEdit.eof then
	%>
<div align="center"><form name="form1" method="post" action="" onSubmit="javascript:{if(this.class_name.value.replace(/\s/gi,'')==''){alert('请输入类别名称');this.class_name.focus();return false;}}">
  <table cellpadding="0" cellspacing="1" border="0" width="95%" align="center" class="table_southidc">
	<tr>
		<td class="back_southidc" height="25" align="center">资料类别添加</td>
	</tr>
	<tr>
	<td>
  <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="1" >
    		  <tr><td colspan="2">

		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="tbbg01">
            <td height="25" align="right" valign="top">类别名称：</td>
            <td align="left" valign="top" class="tbbg01"><input name="class_name_gb" type="text" id="class_name_gb" size="50" maxlength="100" value="<%=rsEdit("class_name_gb")%>">            </td>
          </tr>	         

</table>				
</td></tr>
    <tr bgcolor="#FFFFFF">
      <td height="37" align="right"  class="tbbg01" width="150" >所属分类：</td>
      <td bgcolor="#FFFFFF"  class="tbbg01" width="850">&nbsp;
		 <select name="parent_id" id="parent_id">
                  <%
		dim rs0,sql,selected
		response.write "<option value='0'>（无）作为一级分类</option>"
		sql="select class_id,parent_id,class_name_"&lan_txt&",class_order from Article002class where parent_id=0  order by class_order desc,class_id asc"
		set rs0=server.createobject("adodb.recordset")
		rs0.open sql,conn,1,1
		do while not rs0.eof
			if parent_id_Old=rs0("class_id") then
			selected="selected"
		else
			selected=""
		end if
			response.write "<option value='"&rs0("class_id")&"' "&selected&">"&rs0("class_name_"&lan_txt&"")&"</option>"
			class_id=rs0("class_id")
			call subClass(class_id,parent_id_Old)
		rs0.movenext
		loop
		rs0.close:set rs0=nothing
		%>
      </select></td>
    </tr><!--
    <tr  class="tbbg01">
      <td height="34" align="right" >排序顺序：</td>
      <td >&nbsp;
        <input name="class_order" type="text" id="class_order" size="6" maxlength="6" value="<%=rsEdit("class_order")%>"/>
      <span class="style1">（该值为一个数值，数值越大类别越排前面）        </span></td>
    </tr>-->
    <tr  class="tbbg01">
      <td height="36"  class="tbbg01">&nbsp;</td>
      <td height="36" align="left"  class="tbbg01" >&nbsp;
        <input type="submit" name="Submit" value=" 修 改 " /><input type="hidden" name="action" value="edit" /></td></tr>
  </table>
 </td>
 </tr>
 </table>
</form></div><%
  else
  	response.write "参数传递错误，要修改的类别不存在"
  end if
  rsEdit.close:set rsEdit = nothing
  conn.close:set conn = nothing
  
  %>
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
