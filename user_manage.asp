<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
if session("power")=3 then 
	response.Redirect("user_my_edit.asp")
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
<link rel="stylesheet" href="css/css.css"  type="text/css">
<script language="javascript" src="user_check.js"></script>
</head>
<body>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    
    <td  valign="top">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="35" bgcolor="#ECF5FF">&nbsp;欢迎您，<font color="#FF0000"><%=session("loginusername")%></font>,您现在的身份是<font color="#FF0000"><%if session("power")=1 then
	response.Write("超级管理员")
	elseif session("power")=2 then
	response.Write("普通管理员")
	else
	response.Write("学生")
	end if%></font></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td bgcolor="#99CCFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="30" align="left" bgcolor="#C8EDFB"><table width="60%" border="0" cellspacing="1" cellpadding="1">
				      <tr>
				        <td height="35" align="center" bgcolor="#C8EDFB"><a href="user_add.asp" target="main"> <span > <font color="#000080"><strong>新增用户</strong></font></span></a></td>
				        <td align="center" bgcolor="#C8EDFB"><a href="user_my_edit.asp" target="main"><font color="#000080"> <span ><strong>个人资料</strong></span></font></a></td>
				        <td align="center" bgcolor="#C8EDFB"><a href="user_manage.asp" target="main"><font color="#000080"> <span ><strong>用户管理</strong></span></font></a></td>
				        <td align="center" bgcolor="#C8EDFB"><a href="user_exit.asp" target="main"><font color="#000080"> <span ><strong>安全退出</strong></span></font></a></td>
			          </tr>
			        </table>					</td>
    <td bgcolor="#C8EDFB"></td>
  </tr>
  <tr>
    <td height="30" align="left" bgcolor="#ECF5FF"> 用户管理 &gt;&gt; 用户列表</td>
    <td bgcolor="#ECF5FF"></td>
  </tr>
</table>
  <form action="?action=order" method="post" name="myform" target="_self" id="myform" style="margin:0">
    <table width="100%"  border="0" align="center">
      <tr>
        <td align="right"> 关键字：
          <input name="keyword" type="text" id="keyword" value="<%=keyword%>" />
          <input type="submit" name="Submit2" value="搜索" />
          &nbsp;</td>
      </tr>
    </table>
    <table cellpadding="0" cellspacing="1" border="0" width="100%" align="center" class="table_southidc">
      <tr>
        <td class="back_southidc" height="25" align="center">用户管理</td>
      </tr>
      <tr>
        <td><table width="100%" height="47"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
          <tr align="center" bgcolor="#EEEEEE">
            <td width="9%" height="24">选择</td>
            <td width="11%">用户名</td>
            <td width="8%">姓名</td>
            <td width="10%">联系方式</td>
            <td width="12%">电子邮箱</td>
            <td width="10%">权限</td>
            <td width="20%">备注</td>
            <td width="20%">操作</td>
          </tr>
          <%
	 	dim rs2,sql2
	 	dim rs,sql
		Dim CurrentPage,ii,j,totalnumber
		Const maxperpage=30
		CurrentPage=checkint(Request("page"),1)
		Dim wherestr
		if keyword<>"" then wherestr = wherestr &" and (username like '%"&keyword&"%' or name like '%"&keyword&"%')"
		if session("power")=1 then 
		sql="select * from userinfo where power=2  "&wherestr&" order by id desc"
		elseif session("power")=2 then 
		sql="select * from userinfo where power=3 "&wherestr&" order by id desc"
		end if
		'response.Write(sql)
		set rs =server.CreateObject("adodb.recordset")
		rs.open sql,conn,1,3
		if not rs.eof then
				rs.pagesize=maxperpage
				totalnumber=rs.recordcount
				If CurrentPage>rs.pagecount Then CurrentPage=rs.pagecount
				rs.absolutepage=CurrentPage
				Do while not rs.eof and ii<maxperpage
				ii = ii + 1
	  %>
          <tr bgcolor="#FFFFFF">
            <td align="center"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>" /></td>
            <td align="center" ><font color="blue"><%=rs("username")%></font></td>
            <td align="center" ><%=rs("name")%></td>
            <td align="center" ><%=rs("phone")%>
              </td>
            <td align="center"><%=rs("email")%></td>
            <td align="center"><%if rs("power")= 1 then 
			response.Write("超级管理员")
			elseif rs("power")=2 then
			response.Write("普通管理员")
			else
			response.Write("学生")
			end if%></td>
            <td align="center"><%=rs("content")%></td>
            <td align="center">&nbsp;<a href="user_edit.asp?id=<%=rs("id")%>&page=<%=request("page")%>">修改</a> <a href="user_del.asp?id=<%=rs("id")%>" onClick="return checkDel()">删除</a></td>
          </tr>
          <%
			rs.movenext
		loop
		%>
          <tr bgcolor="#ECF5FF">
            <td height="30" colspan="12" valign="middle" bgcolor="#ECF5FF"><%
		call showpage("user_manage.asp?keyword="&keyword,totalnumber,maxperpage,true,true,"条")
		%></td>
          </tr>
          <%
		else
			response.write "<tr  bgcolor='#FFFFFF'><td height=60 colspan=11>&nbsp;&nbsp;&nbsp;&nbsp;没有数据</td></tr>"
	  	end if
		rs.close
		Set rs =nothing
	  %>
        </table></td>
      </tr>
    </table>
    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="17%"><span style="cursor:hand" onClick="javascript:SelectAll('y')"><font color="#0000FF">全选</font></span><font color="#0000FF"> / <span style="cursor:hand" onClick="javascript:SelectAll('n')">全不选</span></font></td>
        <td width="83%"><input name="Submit1" type="button" id="Submit13" value="删除选中记录" style="border-width:1px;border-color:#999999;border-style:solid; background-color: #eeeeee" onClick="return checkSelectAll(myform,'del','<%=request("lag")%>')" />
          <!--<input name="Submit2" type="button" id="Submit2" value="设为新品" onClick="return checkSelectAll(this.form,'setnew','<%=request("lag")%>')">
      <input name="Submit2" type="button" id="Submit2" value="取消新品" onClick="return checkSelectAll(this.form,'setunnew','<%=request("lag")%>')"></td>
    --></td>
      </tr>
    </table>
  </form></td>

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