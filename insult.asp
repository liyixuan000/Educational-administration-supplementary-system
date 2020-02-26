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
 </table>
<table width="778" align="center" bgcolor="#FFFFFF">
<%if ks=1 then %>
 <tr>
<td id="info" ><%=Description%></td>
  </tr>
 <tr>
  <td colspan="4" style="line-height:20px">
    <div align="center">
    <table width="98%" border="1" cellpadding="3" cellspacing="1" bgcolor="#aec3de" class="yy" bordercolor="#FFFFFF">
     <tr align="center" bgcolor="#ffffff" style="font-weight:bold;line-height:30px;">
	  <td colspan="10">请选择医生：
<%
cid=trim(request("cid"))
sqlClass="select * from Sclass order by id asc"
Set rsClass= Server.CreateObject("ADODB.Recordset")
rsClass.open sqlClass,conn,1,1
if rsClass.eof then 
	response.Write("还没有任何教室，请先添加教室。")
else
	for a=1 to rsClass.recordcount
	b=rsClass("title")
	if rsClass("title")=cid then
	   response.Write("<a href='?cid=" & rsClass("title") & "'><font color='red'>" & rsClass("title") & "</font></a> | ")		
	else
	   response.Write("<a href='?cid=" & rsClass("title") & "'><font color='green'>" & rsClass("title") & "</font></a> | ")
	end if
    rsClass.movenext
	next
end if
rsClass.close
set rsClass=nothing
%>
	  </td>
	  </tr>
     <tr align="center" bgcolor="" style="font-weight:bold;line-height:30px;">
	  <td>日期</td>
	  <td>星期</td>
	  <td width="10%"><font size="1">8:00-9:00</font></td>
	  <td width="10%"><font size="1">9:00-10:00</font></td>
	  <td width="10%"><font size="1">10:00-11:00</font></td>
	  <td width="10%"><font size="1">11:00-12:00</font></td>
	  <td width="10%"><font size="1">14:00-15:00</font></td>
	  <td width="10%"><font size="1">15:00-16:00</font></td>
	  <td width="10%"><font size="1">16:00-17:00</font></td>
	  <td width="10%"><font size="1">17:00-18:00</font></td>
	  
     </tr>
<%for i=-1 to 6
Select Case i
 Case "-1" ys="gray"
 Case "0" ys="red"
 Case "1" ys="black"
end select
dd=DateAdd("d",+i,date())
ww=WeekdayName(Weekday(dd))  
hh=hour(now())
%>
      <tr align='center' bgcolor='#FFFFFF' height="50">
		  <td><font color=<%=ys%>><%=dd%></font></td>
		  <td><%=ww%></td>
		  <%for j=1 to 8%>
<%
sqlY="select * from Info where xx1=#"&dd&"# and xx2='"&ww&"' and xx3='第"&j&"节'"
if cid<>"" then 
   sqlY=sqlY & " and xx4='"&cid&"' "
else
   cid=b
   sqlY=sqlY & " and xx4='"&b&"' "
end if
set rsY=conn.execute(sqlY)
if rsY.eof then
   if i=-1 or (i=0 and hh<=12 and j<=4) or (i=0 and hh>12) then 
	  response.Write("<td>不可约</td>")
   else
	  response.Write("<td><a href='reserve.asp?cid="&cid&"&rq="&dd&"&xq="&ww&"&ks="&j&"'>预约</a></td>")
   end if 
else
   if rsY("sh")<>"待审核" then
	  response.Write("<td bgcolor=#F2FDFF>"&rsY("xx5")&"<br>"&rsY("xx6")&"</td>")
   else
	  response.Write("<td bgcolor=#F2FDFF>已预约</td>")
   end if
end if
rsY.close
Set rsY=nothing
%>
		  <%next%>
       </tr>
<%next%>
      </table>
  	</div>
  </td>
 </tr>
 <%end if%>
</table>
	</td>
  </tr>
  <tr>
    <td colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</div>
</body>
</html>