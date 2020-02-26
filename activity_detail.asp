<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
Dim Id:Id=checkint(request("id"),0)
dim pclass_id:pclass_id=checkint(request("pclass_id"),0)
if Id=0 then jstop("温馨提示：参数错误！")
dim rsEdit
dim rs,sqlstr,Class_id,picurl,postdate,content,title,pagetitle,pclass_id2
sqlstr = "select * from Article where id="&Id
set rsEdit=conn.execute(sqlstr)
class_id =rsEdit("class_id")
picurl=rsEdit("picurl")
postdate=rsEdit("addtimes")
content=rsEdit("content_gb")
title=rsEdit("title_gb")
pagetitle=rsEdit("pagetitle_gb")
pclass_id2=rsEdit("pclass_id")
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
<div align="center"><form action="activity_save.asp?action=Edit" method="post" name="myform" onSubmit="return check();">
<table cellpadding="4" cellspacing="1" class="toptable grid" border="1" width="100%" bordercolor="#FFFFFF">
  <tr>
    <td height="30" align="right" colspan="2">
	<p align="center">活动详情</td>
  </tr>

  <tr>
    <td width="36%" height="30" align="right">活动信息名称：</td>
    <td width="64%" class="category">
        <%=title%></td>
  </tr>

  <tr>
    <td height="30" align="right">所属类别：</td>
    <td class="category">
    <%
	  dim sqlC,class_name,parent_id
	  sqlC="select class_name_"&lan_txt&",parent_id from ArticleClass where  class_id="&class_id&""
	  set rsc=conn.execute (sqlC)
	  if not rsC.eof then
	  	class_name=rsc("class_name_"&lan_txt&"")
	  	response.write class_name
	  end if
	  %></td>
  </tr>
	<tr>
    <td height="30" align="right">活动级别：</td>
    <td class="category"><%
	  'dim sqlC,class_name,parent_id
	  sqlC="select class_name_"&lan_txt&",parent_id from ArticleClass where  class_id="&pclass_id2&""
	  set rsc=conn.execute (sqlC)
	  if not rsC.eof then
	  	class_name=rsc("class_name_"&lan_txt&"")
	  	response.write class_name
	  end if
	  %></td>
  </tr>

  <tr>
    <td height="30" align="right">活动时间：</td>
    <td class="category">
       <%=infotime(postdate)%></td>
  </tr>


  <tr>
    <td height="30" align="right" bgcolor="#99CCFF">活动内容：</td>
    <td class="category"><%=content%></td>
  </tr>


  <tr>
    <td height="30" align="right">附件上传：</td>
    <td class="category">
				  <%if picUrl<>"" then %>
                  <a  href="<%=picurl%>"  target="_blank">下载</a>
                  <%end if%></td>
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
<%
sub classSelectModi(class_id,ArticleClassID)
	dim sql0,rsc0,strSpace,i
	sql0="select class_name_"&lan_txt&",class_id,class_depth from Articleclass where parent_id="&class_id&" "
	set rsc0=conn.execute (sql0)
	do while not rsc0.eof
		strSpace=""
		for i=1 to rsc0("class_depth")
		if i=rsc0("class_depth") then strSpace=strSpace&"├" else strSpace=strSpace&"│"
		next
		if ArticleClassID=rsc0("class_id") then selected="selected" else selected=""
		response.write "<option value='"&rsc0("class_id")&"' "&selected&">"&strSpace&""&rsc0("class_name_"&lan_txt&"")&"</option>"
		class_id=rsc0("class_id")
		call classSelectModi(class_id,ArticleClassID)
		rsc0.movenext
	loop
	rsc0.close:set rsc0=nothing
end sub
%><script>
function check()
{
	if(myform.title_gb.value=='')
	{
		alert("活动名称不能为空！");
		myform.title_gb.focus();
		return false;
	}
	if(myform.postdate.value=='')
	{
		alert("活动时间不能为空！");
		myform.postdate.focus();
		return false;
	}
	if(myform.content_gb.value=='')
	{
		alert("活动内容不能为空！");
		myform.content_gb.focus();
		return false;
	}
}
</script>