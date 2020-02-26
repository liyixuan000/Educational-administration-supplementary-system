<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="inco/conn.asp"-->
<!--#include file="inco/config.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<link rel="stylesheet" href="css/css.css" />
</head>
</head>

<body>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top3.asp"--></td>
  </tr>
<!-- <tr>
    <td width="205" valign="top"> 
        			<li><a href="psy.asp" target="main">
					<span style="background-color: #C0C0C0">
					<font color="#000080">心理健康宣传</font></span></a> 
					<li><a href="psytest.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">心理测试</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">心理咨询预约</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">查看心理咨询师信息</span></font></a></li>

                    <!--<li><a href="#" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">活动认证管理</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动类</span></font></a></li>
					-->
					</td>
    <td width="778" valign="top">
<!--<table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">-->
 <table width="778" align="center" bgcolor="#FAFAFA">
<%if ks=1 then %>
  <tr>
    <td align="center" height="120px" id="link">
<%
   if Session("vote")="yes" or Request.form("xx1")="" then
      response.write "<div>错误：请勿重复提交您的预约信息！<a href='index.asp'>返回首页</a> </div>"
   else
      Set rs = Server.CreateObject( "ADODB.Recordset" )
      sql = "select * from info"
	  rs.open sql,conn,1,3
      rs.addnew
	  for i=1 to 6
	  rs("xx"&i)=Request.form("xx"&i)
	  next
	  rs("sh")="待审核"
	  rs("addtime")=Request.form("addtime")
      rs.update
      rs.close
      Set rs=nothing
      session.TimeOut=6
      response.write"<div>恭喜：您的预约信息提交成功！<p><p><div style=color:#FF0000>请稍后访问查看预约情况！<a href='index.asp'>返回首页</a></div></div>"
   end if%>
	  </td>
    </tr>
<%end if%>
</table>
	</td>
  </tr>
  <tr>
    <td width="778" colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>

</div>
</body>
</html>