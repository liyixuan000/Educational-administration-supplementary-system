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
<title>������������Ϣϵͳ</title>
<link rel="stylesheet" href="css/css.css"  type="text/css">
</head>
<body>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="35" bgcolor="#ECF5FF">&nbsp;��ӭ����<font color="#FF0000"><%=session("loginusername")%></font>,�����ڵ������<font color="#FF0000"><%if session("power")=1 then
	response.Write("��������Ա")
	elseif session("power")=2 then
	response.Write("��ͨ����Ա")
	else
	response.Write("ѧ��")
	end if%></font></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td width="205" valign="top"> <%if session("power")<3 then%>
        			<li><a href="user_add.asp" target="main">
					<span style="background-color: #C0C0C0">
					<font color="#000080">�����û�</font></span></a> </li>
                    <%end if%>
					<li><a href="user_my_edit.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��������</span></font></a></li>
                    <%if session("power")<3 then%>
					<li><a href="user_manage.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">�û�����</span></font></a></li>
                    <%end if%>
                     <li><a href="user_exit.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��ȫ�˳�</span></font></a></li>

					</td>
    <td width="567" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td bgcolor="#99CCFF">
<div align="center"><form action="user_my_save.asp" method="post" name="myform">
<table cellpadding="4" cellspacing="1" class="toptable grid" border="1" width="100%" bordercolor="#FFFFFF">
  <tr>
    <td height="30" align="right" colspan="2">
	<p align="center">�޸��û�(��*�ŵ�Ϊ������)</td>
  </tr>

  <tr>
    <td height="30" align="right">�û�����</td>
    <td class="category">
        <%=username%>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>

  <tr>
    <td height="30" align="right">���룺</td>
    <td class="category">
        <input name="password" type="password" id="password" style="width:300px" size="10">
        &nbsp;<font color="#ff0000">*���޸Ŀ��Բ���</font></td>
  </tr>
	<tr>
    <td height="30" align="right">ȷ�����룺</td>
    <td class="category">
        <input type="password" name="repassword"  style="width:300px"   size="10">
        &nbsp;<font color="#ff0000">*���޸Ŀ��Բ���</font></td>
  </tr>

  <tr>
    <td height="30" align="right">������</td>
    <td class="category">
        <input name="name" type="text" id="name" style="width:300px" value="<%=names%>">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>


  <tr>
    <td height="30" align="right">��ϵ��ʽ��</td>
    <td class="category">
        <input name="phone" type="text" id="phone" style="width:300px" value="<%=phone%>">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>


  <tr>
    <td height="30" align="right">�������䣺</td>
    <td class="category">
        <input name="email" type="text" id="email" style="width:300px" value="<%=email%>"></td>
  </tr>


  <tr>
    <td height="30" align="right">Ȩ�ޣ�</td>
    <td class="category"><select size="1" name="power">
    <%if session("power")=1  then%>
		<option value="1">��������Ա</option>
        <%end if%>
            <%if session("power")=2  then%>
<option value="2">��ͨ����Ա</option>
        <%end if%>
                    <%if session("power")=3  then%>
<option value="2">ѧ��</option>
        <%end if%>
		</select>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>
  <tr>
    <td height="30" align="right">��ע��</td>
    <td class="category"><textarea name="content" rows="6" id="content" style="width:300px"><%=content%></textarea>
        &nbsp;</td>
  </tr>
  <tr>
    <td height="30">��</td>
    <td class="category">
        <input name="submit" type="submit" onClick="return check()" value=" ȷ �� " class="button">
      &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="hidden" name="action" value="edit"><input type="hidden" name="id" value="<%=id%>">
        <input name="reset" type="reset" value=" ������д " class="button">
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
		alert("���벻��Ϊ�գ�");
		myform.password.focus();
		return false;
	}
	if(myform.repassword.value=='')
	{
		alert("ȷ�����벻��Ϊ�գ�");
		myform.repassword.focus();
		return false;
	}
	*/
	if(myform.repassword.value!=myform.password.value)
	{
		alert("������������벻һ�£�");
		myform.repassword.focus();
		return false;
	}
	if(myform.name.value=='')
	{
		alert("��������Ϊ�գ�");
		myform.name.focus();
		return false;
	}
	if(myform.phone.value=='')
	{
		alert("��ϵ��ʽ����Ϊ�գ�");
		myform.phone.focus();
		return false;
	}
}
</script>