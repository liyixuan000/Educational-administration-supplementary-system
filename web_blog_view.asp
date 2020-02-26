<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
<link rel="stylesheet" href="css/css.css" />
<script language="javascript">
<!--function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert(form.elements[i].name + "不能为空!");return false;}	
	if(form.elements[1].value!=form.elements[2].value){
	  alert("验证码错误!");return false;}  
  }
}//-->
</script>
<SCRIPT language=JavaScript>
<!--
var LastCount =0;
function CountStrByte(Message,Total,Used,Remain){ //字节统计
 var ByteCount = 0;
 var StrValue  = Message.value;
 var StrLength = Message.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // 在此判断，减少循环次数
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	Message.value = StrValue.substring(0,i);
			alert("留言内容最多不能超过 " +MaxValue+ " 个字节！\n注意：一个汉字为两字节。");
         ByteCount = MaxValue;
         break;
      }
	}
   Used.value = ByteCount;
   Remain.value = MaxValue - ByteCount;
   LastCount = StrLength;
 }
}
//-->
</SCRIPT>
</head>
<body>
<table width="778" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    <td width="778" valign="top">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" bgcolor="#30ADDB">
  <tr>
    <td bgcolor="#FFFFFF"><!--#include file="Item.asp"-->
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="2">
<%
'显示文章的详细内容,包括文章标题、文章作者、发表时间以及文章内容
id=Request.QueryString("id")
classname=Request.QueryString("classname")
If Not IsNumeric(id) Then
Else
Set rs=conn.Execute("select id,Atitle,Acontent,Aauthor,Adate from tab_article where id="&id&"")
%>
  <tr align="left" bgcolor="#E6F7FE">
    <td colspan="2"><font color="#FF0000"><%=classname%></font> -&gt; <%=rs("Atitle")%></td>
 '   <%'if rs("Annex")<>"" then%>
  </tr>
  <tr>
    <td colspan="2" align="center" class="font1"><%=rs("Atitle")%></td>
  </tr>
  <tr>
    <td width="58%" align="right">操作员：<%=rs("Aauthor")%> </td>
    <td width="42%" align="center"><%=rs("Adate")%></td>
  </tr>
  <tr>
    <td colspan="2" style="line-height:22px">&nbsp;&nbsp;<%=rs("Acontent")%></td>
  </tr>
           </table></td>
  </tr>
 <%
Set rs=Nothing
End IF%> 
</table></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</body>
</html>
