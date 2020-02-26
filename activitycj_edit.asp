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
sqlstr = "select * from apply where id="&Id
set rsEdit=conn.execute(sqlstr)
bnumber =rsEdit("bnumber")
picurl=rsEdit("picurl")
pid=rsEdit("pid")
orders=rsEdit("orders")
prize=rsEdit("prize")
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
                    </table></td>
    <td width="567" valign="top"><table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
      <tr>
        <td nowrap="nowrap" height="500" align="center"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
          <tr>
            <td width="1" height="500" align="center" nowrap="nowrap"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
            </table></td>
            <td align="center" valign="top" height="600" nowrap="nowrap" bgcolor="#99CCFF"><form action="activitycj_save.asp?action=Edit" method="post" name="myform"onsubmit="return check()">
              <div align="center">
                <table cellspacing="0" cellpadding="0" border="1" width="100%" bordercolor="#FFFFFF">
                  <tr>
                    <td nowrap="nowrap" colspan="2" height="30" align="center" bordercolor="#FFFFFF"> 　
                      <p><b><font size="5" face="华文隶书">活动成绩认证申请修改</font></b></p></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" width="30%" bordercolor="#FFFFFF"> 活动名称 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><span class="category">
                      <select name="pid" id="pid">
                        <option value="0">选择活动</option>
                        <%
				dim rsc
			 	set rsc=conn.execute ("select * from Article where 1=1   order by id desc")
				do while not rsc.eof
					if pid=rsc("id") then
					response.write"<option value='"&rsc("id")&"' selected >"&rsc("title_"&lan_txt&"")&"</option>"
					else
					response.write"<option value='"&rsc("id")&"' >"&rsc("title_"&lan_txt&"")&"</option>"
					end if
					'call classSelect(class_id)
					rsc.movenext
				loop
				rsc.close:set rsc=nothing
			 %>
                      </select>
        &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" bordercolor="#FFFFFF"> 活动编号 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input type="text" class="text" name="bnumber" value="<%=bnumber%>" />
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="center" bordercolor="#FFFFFF"> 所获奖项 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input type="text" class="text" name="prize"  value="<%=prize%>" />
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30" align="center" bordercolor="#FFFFFF"> 项目内排名 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input type="text" class="text" name="orders"  value="<%=orders%>"/>
                      填写数字 <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30" align="center" bordercolor="#FFFFFF"> 成绩证明 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><div class="style2" id=picUrl_display style="display:none;"><input name="picUrl" type="text" id="picUrl" size="50" value="<%=picUrl%>"></div>
				  <iframe src="picUpload2.asp" width=100% height=40 frameborder="0" scrolling="no"></iframe>
				  <%if picUrl<>"" then response.Write("<a href='"&picUrl&"' target='_blank' ><font color=red>查看已上传图片</font><font color=blue>"&picUrl&"</font></a>")%></td>
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
	if(myform.bnumber.value=='')
	{
		alert("活动编号不能为空！");
		myform.bnumber.focus();
		return false;
	}
	if(myform.prize.value=='')
	{
		alert("所获奖项不能为空！");
		myform.prize.focus();
		return false;
	}
	if(myform.orders.value=='')
	{
		alert("项目内不能为空！");
		myform.orders.focus();
		return false;
	}
}
</script>