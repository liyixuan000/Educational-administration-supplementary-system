<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim class_id:class_id=checkint(Request("class_id"),0)
sub subClass(class_id)
	dim rsChildSum,counter,sql,countS,strImg,rs,nodeImg,i
	set rsChildSum=server.CreateObject("adodb.recordset")
	rsChildSum.open "select * from Articleclass where parent_ID="&class_ID,conn,1,1
	counter=rsChildSum.recordcount
	rsChildSum.close:set rsChildSum=nothing
	sql="select * from Articleclass where parent_id="&class_id&"   order by class_order desc,class_id asc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	countS=0
	do while not rs.eof
		countS=countS+1
		num=num+1
		strImg=""
		if rs("childSum")>0 then
			nodeImg="<img src='images/treeOpen.gif' align='absmiddle'>"
		else
			nodeImg="<img src='images/treeNode.gif' align='absmiddle'>"
		end if
		for i=1 to rs("class_depth")
			if i>=rs("class_Depth") then 
				if countS>=counter then strImg=strImg&"<img src='images/treeBottom.gif' align='absmiddle'>" else strImg=strImg&"<img src='images/treeMiddle.gif' align='absmiddle'>"
			else
				strImg=strImg&"<img src='images/treeLeft.gif' align='absmiddle'>"
			end if
		next
		response.write "<tr>"

		response.write "<td bgcolor='#ffffff' style='padding-left:6px;' height=20>"&strImg&""&nodeImg&"&nbsp;"&rs("class_name_"&lan_txt&"")&"</td>"
		response.write "<td bgcolor='#ffffff' align=center><input name='class_order' type='text' id='class_order' value='"&rs("class_order")&"' size='6' maxlength='6'><input type='button' name='Submit' value='更新' onClick=""setClassOrders(this.form,"&rs("class_id")&","&num&",'"&request("lag")&"')""></td>"
		response.write "<td bgcolor='#ffffff' align=center><a href=activityclass_add.asp?class_id="&rs("class_id")&"&lag="&request("lag")&">添加子类</a>&nbsp;&nbsp;<a href=activityclass_Edit.asp?class_id="&rs("class_id")&"&parent_id_Old="&rs("parent_id")&">修改</a>&nbsp;&nbsp;<a href='activityclass_del.asp?class_id="&rs("class_id")&"' onclick='return checkDelClass()'>删除</a></td>"
		response.write "</tr>"
		class_id=rs("class_id")
		call subClass(class_id)
		rs.movenext
	loop
	rs.close:set rs=nothing
end sub

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
<link rel="stylesheet" href="css/css.css"  type="text/css">
<script language="javascript" src="act_check.js"></script>
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
</table><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100" height="35" bgcolor="#FFFFFF">&nbsp;<strong>活动分类：<font color="#FF0000"></font></strong></td>
    <td bgcolor="#FFFFFF"><table  border="0" cellspacing="1" cellpadding="1">
      <tr>        <%
				dim rsc,selected
			 	set rsc=conn.execute ("select class_id,class_name_"&lan_txt&",class_depth from Articleclass where parent_id=0  order by class_order desc,class_id asc")
				do while not rsc.eof

			 %>
        <td width="80" height="35" align="center" <%if rsc("class_id")=class_id then 
		response.Write("bgcolor='#1496F5'")
		else
		response.Write("bgcolor='#C8EDFB'")
		end if
		%>
		><a href="?class_id=<%=rsc("class_id")%>"><%=rsc("class_name_"&lan_txt&"")%></a></td>
        <%

					rsc.movenext
				loop
				rsc.close:set rsc=nothing
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
<td bgcolor="#ECF5FF">
<div align="center"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" align="left"> 活动列表 </td>
        <td></td>
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
        <td class="back_southidc" height="25" align="center">活动列表</td>
      </tr>
      <tr>
        <td><table width="100%" height="47"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
          <tr align="center" bgcolor="#EEEEEE">
            <td width="10%" height="24">选择</td>
            <td width="20%">活动名称</td>
            <td width="13%">活动类别</td>
            <td width="15%">活动级别</td>
            <td width="17%">活动时间</td>
            <%if session("power")=2 then%>
            <td width="25%">操作</td>
            <%else%>
            <td width="25%">查看</td>
            <%end if%>
          </tr>
          <%
	 	dim rs2,sql2
	 	dim rs,sql
		Dim CurrentPage,ii,j,totalnumber
		Const maxperpage=30
		CurrentPage=checkint(Request("page"),1)
		
		Dim wherestr
		if keyword<>"" then wherestr = wherestr &" and (title_gb like '%"&keyword&"%' )"
		if class_id>0 then wherestr = wherestr &" and class_id="&class_id
		sql="select * from Article where 1=1  "&wherestr&" order by id desc"
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
            <td align="center" ><a href="activity_detail.asp?id=<%=rs("id")%>&amp;page=<%=request("page")%>"><font color="blue"><%=rs("title_gb")%></font></a></td>
            <td align="center" ><%
	  dim sqlC,class_name,parent_id
	  sqlC="select class_name_"&lan_txt&",parent_id from ArticleClass where  class_id="&rs("class_id")&""
	  set rsc=conn.execute (sqlC)
	  if not rsC.eof then
	  	class_name=rsc("class_name_"&lan_txt&"")
	  	response.write class_name
	  end if
	  %></td>
            <td align="center" ><%
	  'dim sqlC,class_name,parent_id
	  sqlC="select class_name_"&lan_txt&",parent_id from ArticleClass where  class_id="&rs("pclass_id")&""
	  set rsc=conn.execute (sqlC)
	  if not rsC.eof then
	  	class_name=rsc("class_name_"&lan_txt&"")
	  	response.write class_name
	  end if
	  %></td>
            <td align="center"><%=infotime(rs("addtimes"))%></td>
                        <%if session("power")=2 then%>
            <td align="center"><a href="<%=rs("picurl")%>" target="_blank">附件</a> &nbsp;<a href="activity_edit.asp?id=<%=rs("id")%>&amp;page=<%=request("page")%>">修改</a> <a href="activity_del.asp?id=<%=rs("id")%>" onClick="return checkDel()">删除</a></td>
            <%else%>
            <td align="center"><a href="<%=rs("picurl")%>" target="_blank">查看附件</a></td>
            <%end if%>

          </tr>
          <%
			rs.movenext
		loop
		%>
          <tr bgcolor="#ECF5FF">
            <td height="30" colspan="12" valign="middle" bgcolor="#ECF5FF"><%
		call showpage("activity_manage.asp?keyword="&keyword,totalnumber,maxperpage,true,true,"条")
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
    </table><%if session("power")=2 then%>
    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="17%"><span style="cursor:hand" onClick="javascript:SelectAll('y')"><font color="#0000FF">全选</font></span><font color="#0000FF"> / <span style="cursor:hand" onClick="javascript:SelectAll('n')">全不选</span></font></td>
        <td width="83%"><input name="Submit1" type="button" id="Submit13" value="删除选中记录" style="border-width:1px;border-color:#999999;border-style:solid; background-color: #eeeeee" onClick="return checkSelectAll(myform,'del','<%=request("lag")%>')" />
          <!--<input name="Submit2" type="button" id="Submit2" value="设为新品" onClick="return checkSelectAll(this.form,'setnew','<%=request("lag")%>')">
      <input name="Submit2" type="button" id="Submit2" value="取消新品" onClick="return checkSelectAll(this.form,'setunnew','<%=request("lag")%>')"></td>
    --></td>
      </tr>
    </table><%end if%>
  </form>
</div>
</td>

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