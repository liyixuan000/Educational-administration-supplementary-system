<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim pclass_id:pclass_id=checkint(request("pclass_id"),0)
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
<div align="center"><form action="activity_save.asp?action=Add" method="post" name="myform" onSubmit="return check();">
<table cellpadding="4" cellspacing="1" class="toptable grid" border="1" width="100%" bordercolor="#FFFFFF">
  <tr>
    <td height="30" align="right" colspan="2">
	<p align="center">活动添加</td>
  </tr>

  <tr>
    <td height="30" align="right">活动信息名称：</td>
    <td class="category">
        <input name="title_gb" type="text" id="title_gb" style="width:300px">
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>

  <tr>
    <td height="30" align="right">所属类别：</td>
    <td class="category">
      <select name="class_id" id="class_id" onChange="processcityrequest(this.options[this.options.selectedIndex].value)" ><option value="0">选择类别</option>
        <%
				dim rsc,class_id,selected,qqq
				qqq=0
				dim del_class_id
			 	set rsc=conn.execute ("select class_id,class_name_"&lan_txt&",class_depth from Articleclass where parent_id=0  order by class_order desc,class_id asc")
				do while not rsc.eof
					qqq=qqq+1
					if qqq=1 then del_class_id=rsc("class_id")
					class_id=rsc("class_id")
					if pclass_id=rsc("class_id") then selected="selected" else selected=""
					response.write"<option value='"&rsc("class_id")&"' "&selected&">"&rsc("class_name_"&lan_txt&"")&"</option>"
					'call classSelect(class_id)
					rsc.movenext
				loop
				rsc.close:set rsc=nothing
			 %>
      </select>
		<font color="#ff0000">*</font></td>
  </tr>
	<tr>
    <td height="30" align="right">活动级别：</td>
    <td class="category">
        <select name="pclass_id" id="pclass_id"  ><option value="0">选择级别</option>
        <%
				'dim rsc,class_id,selected
			 	set rsc=conn.execute ("select class_id,class_name_"&lan_txt&",class_depth from Articleclass where parent_id="&del_class_id&"  order by class_order desc,class_id asc")
				do while not rsc.eof
					class_id=rsc("class_id")
					if SelectClassId=rsc("class_id") then selected="selected" else selected=""
					response.write"<option value='"&rsc("class_id")&"' "&selected&">"&rsc("class_name_"&lan_txt&"")&"</option>"
					'call classSelect(class_id)
					rsc.movenext
				loop
				rsc.close:set rsc=nothing
			 %>  
      </select>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>

  <tr>
    <td height="30" align="right">活动时间：</td>
    <td class="category">
       <input placeholder="请输入日期" id="dd" class="laydate-icon" name="postdate" onClick="laydate({istime: true, format: 'YYYY-MM-DD hh:mm:ss'})">

        &nbsp;
              
        <font color="#ff0000">*</font></td>
  </tr>


  <tr>
    <td height="30" align="right">活动内容：</td>
    <td class="category"><textarea name="content_gb" rows="6" id="content_gb" style="width:300px"></textarea>
        &nbsp;<font color="#ff0000">*</font></td>
  </tr>


  <tr>
    <td height="30" align="right">附件上传：</td>
    <td class="category">
        <div class="style2" id=picUrl_display style="display:none;"><input name="picUrl" type="text" id="picUrl" size="50"></div>
				  <iframe src="picUpload.asp" width=100% height=40 frameborder="0" scrolling="no"></iframe></td>
  </tr>
  <tr>
    <td height="30">　</td>
    <td class="category">
      <input name="submit" type="submit"  value=" 确 认 " class="button">
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
sub classSelect(class_id)
	dim rsCounter,counter,sql0,rsc0,countS,selected,strSpace,i
	set rsCounter=server.CreateObject("adodb.recordset")
	rsCounter.open "select * from Articleclass where parent_id="&class_id&" ",conn,1,1
	counter=rsCounter.recordcount
	rsCounter.close:set rsCounter=nothing
	sql0="select class_name_"&lan_txt&",class_id,class_depth from Articleclass where parent_id="&class_id&" "
	set rsc0=conn.execute (sql0)
	countS=0
	do while not rsc0.eof
		countS=countS+1
		if request("class_id")=cstr(rsc0("class_id")) or  request("SelectClassId")=cstr(rsc0("class_id")) then
			selected="selected"
		else
			selected=""
		end if
		strSpace=""
		for i=1 to rsc0("class_depth")
			if i>=rsc0("class_depth") then
				if countS>=counter then strSpace=strSpace&"└&nbsp;&nbsp;" else strSpace=strSpace&"├&nbsp;&nbsp;"
			else
				strSpace=strSpace&"│&nbsp;"
			end if
		next
		response.write "<option value='"&rsc0("class_id")&"' "&selected&">"&strSpace&""&rsc0("class_name_"&lan_txt&"")&"</option>"
		class_id=rsc0("class_id")
		call classSelect(class_id)
		rsc0.movenext
	loop
	rsc0.close:set rsc0=nothing
end sub
%>
<script>
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