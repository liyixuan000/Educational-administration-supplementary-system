<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>������������Ϣϵͳ</title>
<link rel="stylesheet" href="css/css.css" />
<script language="javascript">
<!--function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert(form.elements[i].name + "����Ϊ��!");return false;}	
	if(form.elements[1].value!=form.elements[2].value){
	  alert("��֤�����!");return false;}  
  }
}//-->
</script>
<SCRIPT language=JavaScript>
<!--
var LastCount =0;
function CountStrByte(Message,Total,Used,Remain){ //�ֽ�ͳ��
 var ByteCount = 0;
 var StrValue  = Message.value;
 var StrLength = Message.value.length;
 var MaxValue  = Total.value;

 if(LastCount != StrLength) { // �ڴ��жϣ�����ѭ������
	for (i=0;i<StrLength;i++){
		ByteCount  = (StrValue.charCodeAt(i)<=256) ? ByteCount + 1 : ByteCount + 2;
      if (ByteCount>MaxValue) {
      	Message.value = StrValue.substring(0,i);
			alert("����������಻�ܳ��� " +MaxValue+ " ���ֽڣ�\nע�⣺һ������Ϊ���ֽڡ�");
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
'��ʾ���µ���ϸ����,�������±��⡢�������ߡ�����ʱ���Լ���������
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
    <td width="58%" align="right">����Ա��<%=rs("Aauthor")%> </td>
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
