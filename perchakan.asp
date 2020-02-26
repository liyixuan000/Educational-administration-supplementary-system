<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!-- #include file="incc/conn.asp" -->
<!--#include file="incc/callname1.asp"-->


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
    <td colspan="2"><!--#include file="top1.asp"--></td>
  </tr>
  <tr>
    <td width="152" valign="top"> 
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


					</td>
    <td width="620" valign="top">
  <table width="90%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
 <script language=javascript>
function preview() { 
bdhtml=window.document.body.innerHTML;
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr));
window.document.body.innerHTML=prnhtml; 
window.print();
window.document.body.innerHTML=bdhtml; 
         }
</script>
<SCRIPT language=javascript>
//-----------------------------------------------
function checkDate(strDate)
{
	var result = strDate.match(/((^((1[8-9]\d{2})|([2-9]\d{3}))(-)(10|12|0?[13578])(-)(3[01]|[12][0-9]|0?[1-9])$)|(^((1[8-9]\d{2})|([2-9]\d{3}))(-)(11|0?[469])(-)(30|[12][0-9]|0?[1-9])$)|(^((1[8-9]\d{2})|([2-9]\d{3}))(-)(0?2)(-)(2[0-8]|1[0-9]|0?[1-9])$)|(^([2468][048]00)(-)(0?2)(-)(29)$)|(^([3579][26]00)(-)(0?2)(-)(29)$)|(^([1][89][0][48])(-)(0?2)(-)(29)$)|(^([2-9][0-9][0][48])(-)(0?2)(-)(29)$)|(^([1][89][2468][048])(-)(0?2)(-)(29)$)|(^([2-9][0-9][2468][048])(-)(0?2)(-)(29)$)|(^([1][89][13579][26])(-)(0?2)(-)(29)$)|(^([2-9][0-9][13579][26])(-)(0?2)(-)(29)$))/);
	if(result==null)
	{
		return false;
	}
	return true;
}
//-----------------------------------------------

function check()
{
	if (checkDate(document.form2.startdate.value)==false)
	{
		alert("请输入正确的日期格式！");
		return false;
	}
	if (checkDate(document.form2.enddate.value)==false)
	{
		alert("请输入正确的日期格式！");
		return false;
	}
}
</script>
</HEAD>

<BODY>

<script>
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<%  
'取得当前页码
currentpage=request("page")
'response.write currentpage
'response.end
if currentpage<1 or currentpage="" then
  currentpage="1"
end if

'取得搜索关键字
nowstartdate=request("startdate") 
if nowstartdate="" then
  nowstartdate=year(date()-day(date()-1))&"-"&month(date()-day(date()-1))&"-"&day(date()-day(date()-1))
end if
nowenddate=request("enddate") 
if nowenddate="" then
  nowenddate=year(date())&"-"&month(date())&"-"&day(date())
end if  
'nowtype=request("type")

If iscut="1" then
nowisok=request("isok")
End If 

nowEptime_admin=request("Eptime_admin")
nowEptime_wanglai=request("Eptime_wanglai")
nowEptime_xiangmu=request("Eptime_xiangmu")
nowEptime_yuangong=request("Eptime_yuangong")
nowEptime_zhanghu=request("Eptime_zhanghu")
nowbigclass=request("bigclass")
nowsmallclass=request("smallclass")
if nowbigclass="" then nowsmallclass=""
nowkeyword=request("keyword")


dim referer
referer = Request.ServerVariables("HTTP_REFERER")
if request.QueryString("shenhe") = "ok" then
sql1="update Eptime_money set [isok]=true,[shenhe]='"&request.Cookies("eptime_username")&"',[shenhetime]='"&now()&"' where id="&request.QueryString("id")
'这里在审核后更新数据库，审核的字段是isok，审核人，审核时间
conn.execute sql1
response.redirect (referer)
end if
%>
<table width="96%" border="0" cellpadding="0" cellspacing="0" align="center">
<form name="form2">
  <tr> 
    <td width="50" height="50">&nbsp;<img src="images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();"></td>
	<td width="*" align="right">
	
	  &nbsp;<%
	if nowtype="" then
	  sql="select * from Eptime_money_bigclass order by type,id"
	else
	  sql="select * from Eptime_money_bigclass where type="&nowtype&" order by id"
	end if
	set rs_bigclass=conn.execute(sql)
	%>
	  <select name="bigclass" onChange="form2.submit()">
        <option value="">所有大类</option>
        <%
	do while rs_bigclass.eof=false
	%>
        <option value="<%=rs_bigclass("id")%>"<%if trim(cstr(rs_bigclass("id")))=nowbigclass then%> selected="selected"<%end if%>><%=rs_bigclass("bigclass")%></option>
        <%
	  rs_bigclass.movenext
	loop
	%>
      </select>
    <%
	if nowbigclass="" then
	  nowbigclass2=0
	else
	  nowbigclass2=nowbigclass 
	end if
	sql="select * from Eptime_money_smallclass where id_bigclass="&nowbigclass2&" order by id"
	set rs_smallclass=conn.execute(sql)
	%>
	  <select name="smallclass" onChange="form2.submit()">
        <option value="">所有小类</option>
        <%
	do while rs_smallclass.eof=false
	%>
        <option value="<%=rs_smallclass("id")%>"<%if trim(cstr(rs_smallclass("id")))=nowsmallclass then%> selected="selected"<%end if%>><%=rs_smallclass("smallclass")%></option>
        <%
	  rs_smallclass.movenext
	loop
	%>
      </select>
		<%
	sql="select * from Eptime_xiangmu order by id desc"
	set rs_Eptime_xiangmu=conn.execute(sql)
	%>
	  <select name="Eptime_xiangmu" onChange="form2.submit()">
        <option value="">所有业务项目</option>
        <%
	do while rs_Eptime_xiangmu.eof=false
	%>
        <option value="<%=rs_Eptime_xiangmu("id")%>"<%if trim(cstr(rs_Eptime_xiangmu("id")))=nowEptime_xiangmu then%> selected="selected"<%end if%>><%=rs_Eptime_xiangmu("name")%></option>
        <%
	  rs_Eptime_xiangmu.movenext
	loop
	%>
      </select>      <%If request.Cookies("eptime_AdminPower")="2" Then %>
<%else%>  

	<%
	'sql="select * from Eptime_admin where working=1 order by id desc"
	'set rs_Eptime_admin=conn.execute(sql)
	%>
<%End if%>

<%If iscut="1" then%>
	  <select name="isok" onChange="form2.submit()">
	    <option value="">全部状态</option>
        <option value="0"<%if nowisok="0" then%> selected="selected"<%end if%>>
		已审核</option>
        <option value="1"<%if nowisok="1" then%> selected="selected"<%end if%>>
		未审核</option>
      </select>
     
      
<%End if%>
 
	  <br>
		  开始日期：<input name="startdate" value="<%=nowstartdate%>"  style="width:100px" onfocus="javascript:ShowCalendar(this.id)" id="select_date" title="单击选择日期">
	  结束日期：<input name="enddate" value="<%=nowenddate%>" style="width:100px" onfocus="javascript:ShowCalendar(this.id)" id="select_date1" title="单击选择日期">
	  	  <input type="submit" value=" 统计 " class="button" onClick="return check()">&nbsp;
	</td>
  </tr>
</form>  
</table>
<%
If request.Cookies("eptime_AdminPower")="2" Then
sql="select * from Eptime_money where id_login="&request.Cookies("eptime_id")&""
Else
sql="select * from Eptime_money where 1=1"
End if
  


  if nowtype<>"" then
	sql=sql&" and type="&nowtype
  end if 

  if nowisok<>"" Then
  If nowisok=1 then
	sql=sql&" and isok=False"
  ElseIf nowisok=0 Then
    sql=sql&" and isok=true"
  End If 
  end if 

  if nowEptime_admin<>"" then
	sql=sql&" and id_login="&nowEptime_admin
  end if  
    if nowEptime_wanglai<>"" then
	sql=sql&" and wanglai="&nowEptime_wanglai
  end if  

    if nowEptime_zhanghu<>"" then
	sql=sql&" and zhanghu="&nowEptime_zhanghu
  end if  
      if nowEptime_yuangong<>"" then
	sql=sql&" and yuangong="&nowEptime_yuangong
  end if  
      if nowEptime_xiangmu<>"" then
	sql=sql&" and xiangmu="&nowEptime_xiangmu
  end if  


  if nowbigclass<>"" then
	sql=sql&" and id_bigclass="&nowbigclass
  end if
  if nowsmallclass<>"" then
	sql=sql&" and id_smallclass="&nowsmallclass
  end If
  

If(databaseType=0) Then 
    if nowstartdate<>"" then
    sql=sql&" and selldate-#"&nowstartdate&"#>=0"
  end if  
  if nowenddate<>"" then
    sql=sql&" and selldate-#"&nowenddate&"#<=0"
  end if
ElseIf(databaseType=1) Then 
  if nowstartdate<>"" then
    sql=sql&" and selldate-'"&nowstartdate&"'>=0"
  end if  
  if nowenddate<>"" then
    sql=sql&" and selldate-'"&nowenddate&"'<=0"
  end if 
End If 



  if nowkeyword<>"" then
	sql=sql&" and (beizhu like '%"&nowkeyword&"%' or moneyID like '%"&nowkeyword&"%')"
  end if 
  
  if request("order1")<>"" then
    sql=sql&" order by type "&request("order1")
  elseif request("order2")<>"" then
    sql=sql&" order by id_bigclass "&request("order2")
  elseif request("order3")<>"" then
    sql=sql&" order by id_smallclass "&request("order3") 
  elseif request("order4")<>"" then
    sql=sql&" order by login "&request("order4") 
  elseif request("order5")<>"" then
    sql=sql&" order by selldate "&request("order5") 
  elseif request("order6")<>"" then
    sql=sql&" order by type,price "&request("order6") 
  elseif request("order7")<>"" then
    sql=sql&" order by type desc,price "&request("order7")
  elseif request("order8")<>"" then
    sql=sql&" order by id_login "&request("order8")	       
  else
    sql=sql&" order by id desc"  
  end if
  
%>
	<div align="center">
<table width="97%" border="0" cellpadding="0" cellspacing="0" bgcolor="#99CCFF">
<tr>
<td></td>
<td>
<!--startprint-->
<div align="center">
<table cellpadding="4" cellspacing="1" class="toptable grid" border="1" width="100%" bordercolor="#FFFFFF" >
<form name="form1" action="money_del.asp">
  <input type="hidden" name="page" value="<%=currentpage%>">
  <input type="hidden" name="type" value="<%=nowtype%>">
  <input type="hidden" name="bigclass" value="<%=nowbigclass%>">
  <input type="hidden" name="smallclass" value="<%=nowsmallclass%>">
  <input type="hidden" name="startdate" value="<%=nowstartdate%>">
  <input type="hidden" name="enddate" value="<%=nowenddate%>">
  <input type="hidden" name="order1" value="<%=request("order1")%>">
  <input type="hidden" name="order2" value="<%=request("order2")%>">
  <input type="hidden" name="order3" value="<%=request("order3")%>">
  <input type="hidden" name="order4" value="<%=request("order4")%>">
  <input type="hidden" name="order5" value="<%=request("order5")%>">
  <input type="hidden" name="order6" value="<%=request("order6")%>">
  <input type="hidden" name="order7" value="<%=request("order7")%>">
  <input type="hidden" name="order8" value="<%=request("order8")%>">
  <input type="hidden" name="order9" value="<%=request("order9")%>">
  <input type="hidden" name="order10" value="<%=request("order10")%>">
  <input type="hidden" name="order11" value="<%=request("order11")%>">
  <input type="hidden" name="order12" value="<%=request("order12")%>">
  <input type="hidden" name="order13" value="<%=request("order13")%>">
  <input type="hidden" name="order14" value="<%=request("order14")%>">
  <input type="hidden" name="order15" value="<%=request("order15")%>">
  <tr align="center"> 
	<td class="category">
编号</td>
  	<td class="category">
	  <a href="?order5=<%if request("order5")="asc" then%>desc<%else%>asc<%end if%>&page=<%=currentpage%>&type=<%=nowtype%>&Eptime_admin=<%=nowEptime_admin%>&bigclass=<%=nowbigclass%>&smallclass=<%=nowsmallclass%>&startdate=<%=nowstartdate%>&enddate=<%=nowenddate%>" class="title">申请时间<%if request("order5")="asc" then%><img src="images/up2.gif" border="0" hspace="2" align="absmiddle"><%else%><img src="images/down2.gif" border="0" hspace="2" align="absmiddle"><%end if%></a>		
	</td>
	<td class="category">
	  [类型]->[级别]	
	</td>	
<td class="category">说明</td>
<td class="category">申请人</td>
<%If iscut="1" then%>
<td class="category">审核状态</td>
<%End if%>
     </tr>
  <%
  set rs_money =server.createobject("ADODB.RecordSet")	
 ' rs_money.open sql,conn,1,3
  rs_money.Open sql, conn, 3

  if not rs_money.eof then
  'rs_money.pagesize=maxrecord
  rs_money.absolutepage=currentpage
  ii=0
  for currentrec=1 to rs_money.pagesize
    if rs_money.eof then
      exit for
    end if
  %>
  <tr onMouseOver="this.className='highlight'" onMouseOut="this.className=''" onDblClick="javascript:var win=window.open('addchakan_show.asp?id=<%=rs_money("id")%>','帐务详细信息','width=800,height=580,top=20,left=161,toolbar=no,location=no,directories=no,status=no,menubar=no,resizable=no,scrollbars=yes'); win.focus()" <%if ii mod 2 = 0 then%> bgcolor="FFFFFF" <%End if%>>
      <td align="center" <%If rs_money("LastLogin")<>"" Then %>style="color:#FF0000;"<%End If%>><%=rs_money("moneyID")%></td>
    <td align="center" <%If rs_money("LastLogin")<>"" Then %>style="color:#FF0000;"<%End If%>><%=rs_money("selldate")%></td>
	<td align="left">
	  <%
	  sql="select * from Eptime_perClass where id="&rs_money("id_bigclass")
	  set rs_bigclass=conn.execute(sql)
	  %>
	  <%if rs_bigclass.eof=false then%><%=rs_bigclass("bigclass")%><%else%>&nbsp;<%end if%>->
	  <%
	  sql="select * from Eptime_perxingshi where id="&rs_money("id_smallclass")
	  set rs_smallclass=conn.execute(sql)
	  %>
	  <%if rs_smallclass.eof=false then%><%=rs_smallclass("smallclass")%><%else%>&nbsp;<%end if%>
	</td>	

	<td align="center"><%=Left(rs_money("beizhu"),20)%></td>
	<td align="center">
	  <%
	  sql="select * from Eptime_admin where id="&rs_money("id_login")
	  set rs_login=conn.execute(sql)
	  %>
	  <%if rs_login.eof then%><%=rs_money("login")%><%else%><%=rs_login("realname")%><%end if%>	
	</td>

<%If iscut="1" then%>
<td align="center"><%If rs_money("isok")=False then%> <%If request.Cookies("eptime_AdminPower")<>2 then%><a href="?shenhe=ok&id=<%=rs_money("id")%>" title="点击审核">
<%
sql3="select * from Eptime_money where id="&rs_money("id")
  set rs_login=conn.execute(sql3)
  '这两行不懂它的存在

%>
<font color="#CC3300">未审</font></a><%else%><font color="#CC3300">未审</font><%End If %><%else%><a title="审核人：<%=rs_money("shenhe")%>&#13;审核时间：<%=rs_money("shenhetime")%>">已审</a><%End If %></td>
<%End if%>

    </tr>
  <%
  	'end if
    rs_money.movenext
	  ii=ii+1
  Next

  else
  %>
  <tr align="center" onMouseOver="this.className='highlight'" onMouseOut="this.className=''">
    <td <%If iscut="1" then%>colspan="14"<%else%>colspan="13"<%End if%> height="25" align="center" style="color:red"><b>没有找到记录</b></td>
  </tr>
  <%
  end if
  %>
  
  
  
  <%
  if rs_money.recordcount>0 then 
  %> 
  <tr>
    <td <%If iscut="1" then%>colspan="14"<%else%>colspan="13"<%End if%> height="30" class="category">
	<table cellpadding=0 cellspacing=0 width="100%">
	<tr>
	<td width="20%" align="left" style="color:#FF0000;">&nbsp;双击查看详细信息</td>  
	<td width="80%" align="right">
	&nbsp;&nbsp;
      <%=rs_money.recordcount%>&nbsp;条信息&nbsp; 共&nbsp;<%=rs_money.pagecount%>&nbsp;页&nbsp;
  <%
  nowstart=currentpage-3
  if nowstart<1 then
    nowstart=1
  end if
  nowend=currentpage+3
  if nowend>rs_money.pagecount then
    nowend=rs_money.pagecount
  end If
  



  response.write "&nbsp;<a href='?type="&nowtype&"&bigclass="&nowbigclass&"&smallclass="&nowsmallclass&"&Eptime_wanglai="&nowEptime_wanglai&"&Eptime_xiangmu="&nowEptime_xiangmu&"&Eptime_yuangong="&nowEptime_yuangong&"&Eptime_zhanghu="&nowEptime_zhanghu&"&Eptime_admin="&nowEptime_admin&"&isok="&nowisok&"&startdate="&request("startdate")&"&enddate="&nowenddate&"&keyword="&nowkeyword&"&page=1&order1="&request("order1")&"&order2="&request("order2")&"&order3="&request("order3")&"&order4="&request("order4")&"&order5="&request("order5")&"&order6="&request("order6")&"&order7="&request("order7")&"&order8="&request("order8")&"&order9="&request("order9")&"&order10="&request("order10")&"&order11="&request("order11")&"&order12="&request("order12")&"&order13="&request("order13")&"&order14="&request("order14")&"&order15="&request("order15")&"' class='page'>最前页</a>&nbsp;"
  for ipage=nowstart to nowend
    if cstr(ipage)=cstr(currentpage) then
      response.write "&nbsp;<span style='font-weight:bold;color:#5858E6'>" & ipage &"</span>&nbsp;"
    else
      response.write "&nbsp;[&nbsp;<a href='?type="&nowtype&"&bigclass="&nowbigclass&"&smallclass="&nowsmallclass&"&Eptime_wanglai="&nowEptime_wanglai&"&Eptime_xiangmu="&nowEptime_xiangmu&"&Eptime_yuangong="&nowEptime_yuangong&"&Eptime_zhanghu="&nowEptime_zhanghu&"&Eptime_admin="&nowEptime_admin&"&isok="&nowisok&"&startdate="&request("startdate")&"&enddate="&nowenddate&"&keyword="&nowkeyword&"&page="&ipage&"&order1="&request("order1")&"&order2="&request("order2")&"&order3="&request("order3")&"&order4="&request("order4")&"&order5="&request("order5")&"&order6="&request("order6")&"&order7="&request("order7")&"&order8="&request("order8")&"&order9="&request("order9")&"&order10="&request("order10")&"&order11="&request("order11")&"&order12="&request("order12")&"&order13="&request("order13")&"&order14="&request("order14")&"&order15="&request("order15")&"' class='page'>" & ipage &"</a>&nbsp;]&nbsp;"
    end if
  next
  response.write "&nbsp;<a href='?type="&nowtype&"&bigclass="&nowbigclass&"&smallclass="&nowsmallclass&"&Eptime_wanglai="&nowEptime_wanglai&"&Eptime_xiangmu="&nowEptime_xiangmu&"&Eptime_yuangong="&nowEptime_yuangong&"&Eptime_zhanghu="&nowEptime_zhanghu&"&Eptime_admin="&nowEptime_admin&"&isok="&nowisok&"&startdate="&request("startdate")&"&enddate="&nowenddate&"&keyword="&nowkeyword&"&page="&rs_money.pagecount&"&order1="&request("order1")&"&order2="&request("order2")&"&order3="&request("order3")&"&order4="&request("order4")&"&order5="&request("order5")&"&order6="&request("order6")&"&order7="&request("order7")&"&order8="&request("order8")&"&order9="&request("order9")&"&order10="&request("order10")&"&order11="&request("order11")&"&order12="&request("order12")&"&order13="&request("order13")&"&order14="&request("order14")&"&order15="&request("order15")&"' class='page'>最后页</a>&nbsp;"
%>

	</td>
  </tr></table></td></tr>
<%end if%> 
</form>   
</table>
</div>
<!--endprint-->
</td>
<td width="6"></td>
</tr>

</table>

	</div>
<tr>
    <td width="780" valign="bottom" colspan="2"> 
        		<!--#include file="bottom.asp"-->	　</td>
</table>
</div>
</form>	   
   </td>
		    </table>
	</td>
  </tr>
  </table>
</div>
  
        </body>
</html>