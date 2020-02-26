<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
<link rel="stylesheet" href="css/css.css" />
</head>
<body>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top2.asp"--></td>
  </tr>
  <tr>
    <td width="205" valign="top"> 
        			<li><a href="data.asp" target="main">
					<span style="background-color: #C0C0C0">
					<font color="#000080">最新资料</font></span></a> 
					<li><a href="per_add.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">发布个人信息</span></font></a></li>
					<li><a href="perchakan.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">个人信息发布列表</span></font></a></li>
					<li><a href="data_add.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">资料信息上传</span></font></a></li>
					<li><a href="datachakan.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">资料信息上传列表</span></font></a></li>
                    <!--<li><a href="#" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">活动认证管理</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动类</span></font></a></li>
					-->
					</td>
    <td width="567" valign="top">
  <table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
  <tr>
        <td nowrap height="500" align="center">
            <table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
                <tr>
                    <td width="1" height="500" align="center" nowrap>
                    	<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
													</table>
                    </td>
					<td align="center" valign="top" height="600" nowrap bgcolor="#99CCFF">
<form name="AddEquip" method="post" action="per_deal.asp"  enctype="multipart/form-data" onsubmit="return Check()">
<div align="center">
<table cellspacing="0" cellpadding="0" border="1" width="100%" bordercolor="#FFFFFF">
    <tr>
        <td nowrap colspan="2" height="30" align="center" bordercolor="#FFFFFF">
       　<p><b><font size="5" face="华文隶书">个人信息发布</font></b>        </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" width="30%" bordercolor="#FFFFFF">
        学号</td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqName"></td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" width="30%" bordercolor="#FFFFFF">
        联系方式
        </td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqPhone">
        </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
        资料类别
        </td>
		<td nowrap bordercolor="#FFFFFF">
        &nbsp;
        <select size="1" name="EqProo">
		<option selected>考研资料</option>
		<option>出国留学</option>
		<option>就业信息</option>
		</select>
        </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
       资料名称</td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqAmount">
               </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
       资料信息简介</td>
	  <td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqPrice">
                </td>
    </tr>
	<tr>
		<td colspan="2" align="center" height="30" bordercolor="#FFFFFF">
		<input type="submit" name="Post" value="提交" class="button">
		<input type="reset" name="重置" value="重置" class="button">
		<input type="button" name="返回" value="返回" class="button" onclick="Javascript:history.go(-1)">
		</td>
	</tr>
</table>
</div>
</form>	   
   </td>
		  </tr>
</table>
	</td>
  </tr>
  </table>
</div>
  <tr>
    <td width="772" valign="top" colspan="2"> 
        		<!--#include file="bottom.asp"-->	　</td>
    </body>
</html>