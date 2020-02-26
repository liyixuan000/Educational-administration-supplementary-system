<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim class_id:class_id=checkint(Request("class_id"),0)
dim pid:pid=checkint(Request("pid"),0)

dim isok:isok=checkint(Request("isok"),0)
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
<script language="javascript" src="photo_check.js"></script><script type="text/javascript" src="js/laydate.js"></script>

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
<body><script src="area_Ajax.js" ></script> <script language=javascript>
function preview() { 
	htmlcode=window.document.body.innerHTML; 
	sprnstr="<!--startprint-->"; 
	eprnstr="<!--endprint-->"; 
	var prnhtml=htmlcode.substr(htmlcode.indexOf(sprnstr)+17); 
	prnhtml=""+prnhtml.substr(prnhtml.indexOf(sprnstr)+17);
	prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr))+""; 
	window.document.body.innerHTML=prnhtml; 
	window.title="";
	window.print(); 
	window.document.body.innerHTML=htmlcode;
}
</script>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top3.asp"--><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="35" bgcolor="#ECF5FF">&nbsp;欢迎您，<font color="#FF0000"><%=session("loginusername")%></font>,您现在的身份是<font color="#FF0000"><%if session("power")=1 then
	response.Write("超级管理员")
	elseif session("power")=2 then
	response.Write("普通管理员")
	else
	response.Write("学生")
	end if%></font></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td width="120" valign="top"><table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="photo_manage.asp" target="main"> <span > <font color="#fff">心理咨询师信息列表</font></span></a></td>
      </tr>
      <%if session("power")=2 then%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="photo_add.asp" target="main"><font color="#fff"> <span >添加心理咨询师</span></font></a>
          </li></td>
      </tr>
      <%end if%>
      <%if session("power")=2 then%>
      <%end if%>
    </table></td>
    <td width="567" valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td bgcolor="#ECF5FF">
<div align="center"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" align="left"> 心理咨询师信息列表 </td>
        <td></td>
      </tr>
  </table>
  <form action="?action=order" method="post" name="myform" target="_self" id="myform" style="margin:0"><!--startprint-->
    <table cellpadding="0" cellspacing="1" border="0" width="100%" align="center" class="table_southidc">
      <tr>
        <td class="back_southidc" height="25" align="center">心理咨询师信息列表</td>
      </tr>
      <tr>
        <td><table width="100%" height="47"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
          <tr align="center" bgcolor="#EEEEEE">
            <td width="8%" height="24" bgcolor="#EEEEEE">选择</td>
            
            <td width="11%" bgcolor="#EEEEEE">姓名</td><td width="13%" bgcolor="#EEEEEE">从业资格证编号</td>
            <td width="11%" bgcolor="#EEEEEE">照片</td>
            
            <%if session("power")=3 then%>
            <%else%>
            <td width="23%" bgcolor="#EEEEEE">操作</td>
            <%end if%>
          </tr>
          <%
	 	dim rs2,sql2
	 	dim rs,sql
		Dim CurrentPage,ii,j,totalnumber
		Const maxperpage=30
		CurrentPage=checkint(Request("page"),1)
		
		Dim wherestr
		if keyword<>"" then wherestr = wherestr &" and (bnumber like '%"&keyword&"%' )"
		if saferequest("starttime")<>"" and saferequest("endtime")<>"" then wherestr=wherestr &" and (addtimes between #"&saferequest("starttime")&"# and #"&saferequest("endtime")&"# )"
		if pid>0 then wherestr = wherestr &" and pid="&pid
		if request("isok")="a" then
		 
		 else
		 	if request("isok")<>"" then wherestr = wherestr &" and isok="&isok
		 	
		 end if

				sql="select * from photo where 1=1  "&wherestr&" order by addtimes desc "
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
            <td align="center" bgcolor="#FFFFFF"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>" /></td>
            <td align="center" bgcolor="#FFFFFF" ><a href="photo_detail.asp?id=<%=rs("id")%>&amp;page=<%=request("page")%>"><font color="blue"><%=rs("bnumber")%></font></a></td>
            <td align="center" bgcolor="#FFFFFF" ><%=rs("prize")%></td>
            <td align="center" bgcolor="#FFFFFF" ><%if rs("picurl")<>"" then %>
           <a href="<%=rs("picurl")%>" target="_blank"> <img src="<%=rs("picurl")%>"  width="100" /></a>
            <%end if%></td>
            <%if session("power")=3 then%>
         <%else%>
            <td align="center" bgcolor="#FFFFFF">&nbsp;<a href="photo_edit.asp?id=<%=rs("id")%>&amp;page=<%=request("page")%>">修改</a> <a href="photo_del.asp?id=<%=rs("id")%>" onClick="return checkDel()">删除</a></td>
            <%end if%>
           

            


          </tr>
          <%
			rs.movenext
		loop
		%>
          <tr bgcolor="#ECF5FF">
            <td height="30" colspan="13" valign="middle" bgcolor="#ECF5FF"><%
		call showpage("photo_manage.asp?keyword="&keyword,totalnumber,maxperpage,true,true,"条")
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
    </table><!--endprint-->
    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="17%"><%if session("power")=3 then%><span style="cursor:hand" onClick="javascript:SelectAll('y')"><font color="#0000FF">全选</font></span><font color="#0000FF"> / <span style="cursor:hand" onClick="javascript:SelectAll('n')">全不选</span></font></td>
        <td width="83%"><input name="Submit1" type="button" id="Submit13" value="删除选中记录" style="border-width:1px;border-color:#999999;border-style:solid; background-color: #eeeeee" onClick="return checkSelectAll(myform,'del','<%=request("lag")%>')" />
          <!--<input name="Submit2" type="button" id="Submit2" value="设为新品" onClick="return checkSelectAll(this.form,'setnew','<%=request("lag")%>')">
      <input name="Submit2" type="button" id="Submit2" value="取消新品" onClick="return checkSelectAll(this.form,'setunnew','<%=request("lag")%>')"></td>
    --><%end if%></td>
      </tr>
    </table>
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