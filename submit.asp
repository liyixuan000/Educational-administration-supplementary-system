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
					<font color="#000080">����������</font></span></a> 
					<li><a href="psytest.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">�������</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">������ѯԤԼ</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">�鿴������ѯʦ��Ϣ</span></font></a></li>

                    <!--<li><a href="#" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">���֤����</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��ӻ</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��ӻ��</span></font></a></li>
					-->
					</td>
    <td width="778" valign="top">
<!--<table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">-->
 <table width="778" align="center" bgcolor="#FAFAFA">
<%if ks=1 then %>
  <tr>
    <td align="center" height="120px" id="link">
<%
   if Session("vote")="yes" or Request.form("xx1")="" then
      response.write "<div>���������ظ��ύ����ԤԼ��Ϣ��<a href='index.asp'>������ҳ</a> </div>"
   else
      Set rs = Server.CreateObject( "ADODB.Recordset" )
      sql = "select * from info"
	  rs.open sql,conn,1,3
      rs.addnew
	  for i=1 to 6
	  rs("xx"&i)=Request.form("xx"&i)
	  next
	  rs("sh")="�����"
	  rs("addtime")=Request.form("addtime")
      rs.update
      rs.close
      Set rs=nothing
      session.TimeOut=6
      response.write"<div>��ϲ������ԤԼ��Ϣ�ύ�ɹ���<p><p><div style=color:#FF0000>���Ժ���ʲ鿴ԤԼ�����<a href='index.asp'>������ҳ</a></div></div>"
   end if%>
	  </td>
    </tr>
<%end if%>
</table>
	</td>
  </tr>
  <tr>
    <td width="778" colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>

</div>
</body>
</html>