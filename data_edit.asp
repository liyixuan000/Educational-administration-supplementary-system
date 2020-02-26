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
sqlstr = "select * from Article002 where id="&Id
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
<body><script src="data_Ajax.js" ></script>
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
<td bgcolor="#99CCFF">
<div align="center"><form action="data_save.asp?action=Edit" method="post" name="myform" onSubmit="return check();">
<table cellpadding="4" cellspacing="1" class="toptable grid" border="1" width="100%" bordercolor="#FFFFFF">
  <tr>
    <td height="30" align="right" colspan="2">
	<p align="center">资料修改</td>
  </tr>

  <tr>
    <td height="30" align="right">资料名称：</td>
    <td class="category">
        <input name="title_gb" type="text" id="title_gb" style="width:300px" value="<%=title%>">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>

  <tr>
    <td height="30" align="right">所属类别：</td>
    <td class="category">
    <select name="class_id" id="class_id" onChange="processcityrequest(this.options[this.options.selectedIndex].value)"><option value="0">选择类别</option>
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
		<font color="#ff0000">*</font></td>
  </tr>
	<tr>
    <td height="30" align="right">资料形式：</td>
    <td class="category">
                <select name="pclass_id" id="pclass_id"  >
<script  language="javascript">processcityrequest(<%=class_id%>,<%=pclass_id2%>);</script>   
      </select>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>
<!--
  <tr>
    <td height="30" align="right">资料时间：</td>
    <td class="category">
       <input placeholder="请输入日期" id="dd" class="laydate-icon" name="postdate" value="<%=infotime(postdate)%>" onClick="laydate({istime: true, format: 'YYYY-MM-DD hh:mm:ss'})">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>
-->

  <tr>
    <td height="30" align="right" bgcolor="#99CCFF">资料说明：</td>
    <td class="category"><textarea name="content_gb" rows="6" id="content_gb" style="width:300px"><%=content%></textarea>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>


  <tr>
    <td height="30" align="right">附件上传：</td>
    <td class="category">
        <div class="style2" id=picUrl_display style="display:none;"><input name="picUrl" type="text" id="picUrl" size="50" value="<%=picUrl%>"></div>
				  <iframe src="picUpload.asp" width=100% height=40 frameborder="0" scrolling="no"></iframe>
				  <%if picUrl<>"" then response.Write("<a href='"&picUrl&"' target='_blank' ><font color=red>查看已上传图片</font><font color=blue>"&picUrl&"</font></a>")%></td>
  </tr>
  <tr>
    <td height="30">　</td>
    <td class="category">
      <input name="submit" type="submit"  value=" 确 认 " class="button"><input name="id" type="hidden" id="id" value="<%=request("id")%>" />
      &nbsp;&nbsp;&nbsp;&nbsp;

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
<%
sub classSelectModi(class_id,Article002ClassID)
	dim sql0,rsc0,strSpace,i
	sql0="select class_name_"&lan_txt&",class_id,class_depth from Article002class where parent_id="&class_id&" "
	set rsc0=conn.execute (sql0)
	do while not rsc0.eof
		strSpace=""
		for i=1 to rsc0("class_depth")
		if i=rsc0("class_depth") then strSpace=strSpace&"├" else strSpace=strSpace&"│"
		next
		if Article002ClassID=rsc0("class_id") then selected="selected" else selected=""
		response.write "<option value='"&rsc0("class_id")&"' "&selected&">"&strSpace&""&rsc0("class_name_"&lan_txt&"")&"</option>"
		class_id=rsc0("class_id")
		call classSelectModi(class_id,Article002ClassID)
		rsc0.movenext
	loop
	rsc0.close:set rsc0=nothing
end sub
%><script>
function check()
{
	if(myform.title_gb.value=='')
	{
		alert("资料名称不能为空！");
		myform.title_gb.focus();
		return false;
	}
	if(myform.postdate.value=='')
	{
		alert("资料时间不能为空！");
		myform.postdate.focus();
		return false;
	}
	if(myform.content_gb.value=='')
	{
		alert("资料内容不能为空！");
		myform.content_gb.focus();
		return false;
	}
}
</script>