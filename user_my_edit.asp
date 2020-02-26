<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim rsEdit
dim rs,sqlstr,names,phone,email,content,power,username,id
sqlstr = "select top 1 * from userinfo where username='"&session("loginusername")&"'"
set rsEdit=conn.execute(sqlstr)
id=rsEdit("id")
username =rsEdit("username")
power=rsEdit("power")
email=rsEdit("email")
content=rsEdit("content")
phone=rsEdit("phone")
names=rsEdit("name")
rsEdit.close
set rsEdit=nothing
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
</table></td>
  </tr>
  <tr>
    <td width="205" valign="top"> <%if session("power")<3 then%>
        			<li><a href="user_add.asp" target="main">
					<span style="background-color: #C0C0C0">
					<font color="#000080">新增用户</font></span></a> </li>
                    <%end if%>
					<li><a href="user_my_edit.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">个人资料</span></font></a></li>
                    <%if session("power")<3 then%>
					<li><a href="user_manage.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">用户管理</span></font></a></li>
                    <%end if%>
                     <li><a href="user_exit.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">安全退出</span></font></a></li>

					</td>
    <td width="567" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td bgcolor="#99CCFF">
<div align="center"><form action="user_my_save.asp" method="post" name="myform">
<table cellpadding="4" cellspacing="1" class="toptable grid" border="1" width="100%" bordercolor="#FFFFFF">
  <tr>
    <td height="30" align="right" colspan="2">
	<p align="center">修改用户(带*号的为必填项)</td>
  </tr>

  <tr>
    <td height="30" align="right">用户名：</td>
    <td class="category">
        <%=username%>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>

  <tr>
    <td height="30" align="right">密码：</td>
    <td class="category">
        <input name="password" type="password" id="password" style="width:300px" size="10">
        &nbsp;<font color="#ff0000">*不修改可以不填</font></td>
  </tr>
	<tr>
    <td height="30" align="right">确认密码：</td>
    <td class="category">
        <input type="password" name="repassword"  style="width:300px"   size="10">
        &nbsp;<font color="#ff0000">*不修改可以不填</font></td>
  </tr>

  <tr>
    <td height="30" align="right">姓名：</td>
    <td class="category">
        <input name="name" type="text" id="name" style="width:300px" value="<%=names%>">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>


  <tr>
    <td height="30" align="right">联系方式：</td>
    <td class="category">
        <input name="phone" type="text" id="phone" style="width:300px" value="<%=phone%>">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>


  <tr>
    <td height="30" align="right">电子邮箱：</td>
    <td class="category">
        <input name="email" type="text" id="email" style="width:300px" value="<%=email%>"></td>
  </tr>


  <tr>
    <td height="30" align="right">权限：</td>
    <td class="category"><select size="1" name="power">
    <%if session("power")=1  then%>
		<option value="1">超级管理员</option>
        <%end if%>
            <%if session("power")=2  then%>
<option value="2">普通管理员</option>
        <%end if%>
                    <%if session("power")=3  then%>
<option value="2">学生</option>
        <%end if%>
		</select>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>
  <tr>
    <td height="30" align="right">备注：</td>
    <td class="category"><textarea name="content" rows="6" id="content" style="width:300px"><%=content%></textarea>
        &nbsp;</td>
  </tr>
  <tr>
    <td height="30">　</td>
    <td class="category">
        <input name="submit" type="submit" onClick="return check()" value=" 确 认 " class="button">
      &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="hidden" name="action" value="edit"><input type="hidden" name="id" value="<%=id%>">
        <input name="reset" type="reset" value=" 重新填写 " class="button">
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
<script>
function check()
{

	/*
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
	*/
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