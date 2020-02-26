<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
Dim Id:Id=checkint(request("id"),0)
dim pclass_id:pclass_id=checkint(request("pclass_id"),0)
if Id=0 then jstop("温馨提示：参数错误！")
dim rsEdit
dim rs,sqlstr,Class_id,picurl,postdate,content,title,pagetitle,pclass_id2,bnumber,orders,prize,pid
sqlstr = "select * from person where id="&Id
set rsEdit=conn.execute(sqlstr)
bnumber =rsEdit("bnumber")
picurl=rsEdit("picurl")
pid=rsEdit("pid")
orders=rsEdit("orders")
prize=rsEdit("prize")
title=rsEdit("title")
content=rsEdit("content")
Class_id=rsEdit("Class_id")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
<link rel="stylesheet" href="css/css.css"  type="text/css"><script type="text/javascript" src="js/laydate.js"></script>

<style type="text/css">
*{margin:0;padding:0;list-style:none;}
html{background-color:#E3E3E3; font-size:14px; color:#000; font-family:'微软雅黑'}
h2{line-height:30px; font-size:20px;}
a,a:hover{ text-decoration:none;}
pre{font-family:'微软雅黑'}
.box{width:970px; padding:10px 20px; background-color:#fff; margin:10px auto;}
.box a{padding-right:20px;}
.demo1,.demo2,.demo3,.demo4,.demo5,.demo6{margin:25px 0;}
h3{margin:10px 0;}
.layinput{height: 22px;line-height: 22px;width: 150px;margin: 0;}
</style>
</head>
<body><script src="area_Ajax.js" ></script>
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
    <td width="120" valign="top"><table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
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
    </table></td>
    <td width="567" valign="top"><table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
      <tr>
        <td nowrap="nowrap" height="500" align="center"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
          <tr>
            <td width="1" height="500" align="center" nowrap="nowrap"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
            </table></td>
            <td align="center" valign="top" height="600" nowrap="nowrap" bgcolor="#99CCFF"><form action="person_save.asp?action=Edit" method="post" name="myform"onsubmit="return check()">
              <div align="center">
                <table cellspacing="0" cellpadding="0" border="1" width="100%" bordercolor="#FFFFFF">
                  <tr>
                    <td nowrap="nowrap" colspan="2" height="30" align="center" bordercolor="#FFFFFF"> 　
                      <p><b><font size="5" face="华文隶书">个人信息发布修改</font></b></p></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" width="30%" bordercolor="#FFFFFF"> 资料类别 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><span class="category">
                      <select name="class_id" id="class_id"><option value="0">选择类别</option>
		     <%
				dim rsc,selected
			 	set rsc=conn.execute ("select class_id,class_name_"&lan_txt&",class_depth from Article002class where parent_id=0  order by class_order desc,class_id asc")
				do while not rsc.eof
					'class_id=rsc("class_id")
					if class_id=rsc("class_id") then selected="selected" else selected=""
					response.write"<option value='"&rsc("class_id")&"' "&selected&">"&rsc("class_name_"&lan_txt&"")&"</option>"
					'call classSelectModi(rsc("class_id"),class_id)
					rsc.movenext
				loop
				rsc.close:set rsc=nothing
			 %>
      </select>
		<font color="#ff0000">*</font></font></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" bordercolor="#FFFFFF"> 学号 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input type="text" class="text" name="bnumber" value="<%=bnumber%>" />
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" bordercolor="#FFFFFF"> 联系方式 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input type="text" class="text" name="prize"  value="<%=prize%>" />
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" bordercolor="#FFFFFF"> 资料名称 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input name="title" type="text" class="text" id="title" value="<%=title%>" />
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" bordercolor="#FFFFFF"> 资料信息简介 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><textarea name="content" rows="6" class="text" id="content"><%=content%></textarea>
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                  <tr>
                    <td colspan="2" align="center" height="30" bordercolor="#FFFFFF"><input type="submit" name="Post" value="提交" class="button" />
                      <input type="reset" name="重置" value="重置" class="button" /><input name="id" type="hidden" id="id" value="<%=request("id")%>" />
                      <input type="button" name="返回" value="返回" class="button" onClick="Javascript:history.go(-1)" /></td>
                  </tr>
                </table>
              </div>
            </form></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
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
	if(myform.class_id.value=='0')
	{
		alert("资料类别不能为空！");
		myform.class_id.focus();
		return false;
	}
	if(myform.bnumber.value=='')
	{
		alert("学号不能为空！");
		myform.bnumber.focus();
		return false;
	}
	if(myform.prize.value=='')
	{
		alert("联系方式不能为空！");
		myform.prize.focus();
		return false;
	}
	if(myform.title.value=='')
	{
		alert("资料名称不能为空！");
		myform.title.focus();
		return false;
	}
	if(myform.content.value=='')
	{
		alert("资料信息简介不能为空！");
		myform.content.focus();
		return false;
	}
}
</script>