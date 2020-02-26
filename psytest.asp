<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="include/config1.asp"-->
<!--#include file="include/conn1.asp"-->

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

                    <<li><a href="#" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">活动认证管理</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动类</span></font></a></li>
					
					</td>-->

    <td width="778" valign="top">
<table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
<!--这只是一个标记-->
</script>
<table cellSpacing=0 cellPadding=0 align=center background="images/topbj.jpg" width="778">
  <!--<tr>
    <td colSpan=3 id="banner"><%=SiteName%></td>-->
  </tr>
  <tr>
    <td colSpan=3 id="menu" bgcolor="#0099FF">抑郁症自测量</td>
  </tr>
</table>
<div align="center">
<table width="100%" bgcolor="#FAFAFA" border="1" bordercolor="#FFFFFF">
<form id="form1" name="fm" method="post" action="#" onsubmit="return checkForm(this)"> 
  <tr>
   <td colspan="4" id="BigTitle" bgcolor="#99CCFF">一、选择题 <%=ChoiceDes%> </td>
 </tr>
<%
 NumC=20
 'ChoiceNum
 sqlChoice="select top "&NumC&" * from Choice order by id asc"
 set rsChoice=server.createobject("adodb.recordset") 
 rsChoice.open sqlChoice,conn,1,1
 Do while not rsChoice.eof
 For k = 1 to NumC
%>
 <tr onMouseOver="this.bgColor='#F0F7FF'" onMouseOut="this.bgColor='#FAFAFA'">
  <td colspan="4" bgcolor="#99CCFF"><table width="98%" cellpadding="0" cellspacing="2" align="center">
    <!--<tr>
      <td colspan="2"><b><%=k%>.<%=rsChoice("Topic")%></b></td>
    </tr>-->
    <tr>
      <td width="50%"><input type="radio" name="Choice<%=k%>" value="A" title="请选择单选“第<%=k%>题”~!" />
        A、<%=rsChoice("A")%></td>
      <td><input type="radio" name="Choice<%=k%>" value="B" title="请选择单选“第<%=k%>题”~!" />
        B、<%=rsChoice("B")%></td>
    </tr>
    <tr>
      <td width="50%"><input type="radio" name="Choice<%=k%>" value="C" title="请选择单选“第<%=k%>题”~!" />
        C、<%=rsChoice("C")%></td>
      <td><input type="radio" name="Choice<%=k%>" value="D" title="请选择单选“第<%=k%>题”~!" />
        D、<%=rsChoice("D")%></td>
    </tr>
   </table>
  </td>
 </tr>
<%
rsChoice.movenext 
Next
Loop
rsChoice.close
set rsChoice=nothing
%>

'<%if dx<>0 then %> 
<!--<tr>
   <td colspan="4" id="BigTitle" bgcolor="#99CCFF"><br>二、多选题 <%=MChoiceDes%> </td>
 </tr>
<%
 NumM=MChoiceNum
 sqlMChoice="select top "&NumM&" * from MultiChoice order by id asc"
 set rsMChoice=server.createobject("adodb.recordset") 
 rsMChoice.open sqlMChoice,conn,1,1
 Do while not rsMChoice.eof
 For m = 1 to NumM
%>
 <tr onMouseOver="this.bgColor='#F0F7FF'" onMouseOut="this.bgColor='#FAFAFA'">
  <td colspan="4" bgcolor="#99CCFF"><table width="98%" cellpadding="0" cellspacing="2" align="center">
    <tr>
      <td colspan="2"><b><%=m%>.<%=rsMChoice("Topic")%></b></td>
    </tr>
    <tr>
      <td width="50%"><input type="checkbox" name="MChoice<%=m%>" value="A" title="请选择多选“第<%=m%>题”~~min:1" />
        A、<%=rsMChoice("A")%></td>
      <td><input type="checkbox" name="MChoice<%=m%>" value="B" title="请选择多选“第<%=m%>题”~~min:1" />
        B、<%=rsMChoice("B")%></td>
    </tr>
    <tr>
      <td width="50%"><input type="checkbox" name="MChoice<%=m%>" value="C" title="请选择多选“第<%=m%>题”~~min:1" />
        C、<%=rsMChoice("C")%></td>
      <td><input type="checkbox" name="MChoice<%=m%>" value="D" title="请选择多选“第<%=m%>题”~~min:1" />
        D、<%=rsMChoice("D")%></td>
    </tr>
   </table>
  </td>
 </tr>
<%
rsMChoice.movenext 
Next
Loop
rsMChoice.close
set rsMChoice=nothing
end if
%>

<%if pd<>0 then %> 
<tr>
   <td colspan="4" id="BigTitle" bgcolor="#99CCFF"><br>三、判读题 <%=JudgeDes%> </td>
 </tr>
<%
 NumJ=JudgeNum
 sqlJudge="select top "&NumJ&" * from Judge order by id asc"
 set rsJudge=server.createobject("adodb.recordset") 
 rsJudge.open sqlJudge,conn,1,1
 Do while not rsJudge.eof
 For w = 1 to NumJ
%>
 <tr onMouseOver="this.bgColor='#F0F7FF'" onMouseOut="this.bgColor='#FAFAFA'">
  <td colspan="4" bgcolor="#99CCFF"><table width="98%" cellpadding="0" cellspacing="2" align="center">
    <tr>
      <td colspan="2" width="80%"><b><%=w%>.<%=rsJudge("Topic")%></b></td>
	  <td>
	  <input type="radio" name="Judge<%=w%>" value="A" title="请选择判断“第<%=w%>题”~!" />对 
	  <input type="radio" name="Judge<%=w%>" value="B" title="请选择判断“第<%=w%>题”~!" />错</td>
    </tr>
   </table>
  </td>
 </tr>
<%
rsJudge.movenext 
Next
Loop
rsJudge.close
set rsJudge=nothing
end if
%>
-->

  <tr>
    <td colspan="5" align="center" bgcolor="#99CCFF">
	<input type="submit" name="Submit" value="提交测试结果" onclick="return checkClass();" style="width:115; height:42; font-size:18px"/></td>
  </tr>
</form>
</table>
</div>
<%
conn.close
set conn=nothing
%>
<!--这也是一个标记-->
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