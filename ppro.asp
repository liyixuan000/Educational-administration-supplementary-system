<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>������������Ϣϵͳ</title>
<link rel="stylesheet" href="css/css.css" />
</head>
<body>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    <td width="205" valign="top"> 
        			<li><a href="social.asp" target="main">
					<span style="background-color: #C0C0C0">
					<font color="#000080">�������ʵ����Ŀ</font></span></a> 
					<li><a href="ppro.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">�����������</span></font></a></li>
					<li><a href="pprochakan" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">�����������״̬�鿴</span></font></a></li>
					<li><a href="pro.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">���ʵ����֤�ύ</span></font></a></li>
					<li><a href="pro_deal.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��֤״̬�鿴</span></font></a></li>

					<li><a href="honor.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">���ʵ����������</span></font></a></li>
					<li><a href="honorchakan.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��������״̬�鿴</span></font></a></li>

                    <li><a href="project.asp" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">������Ŀ��Ϣ�ϴ�</span></font></a></li>

                    <!--<li><a href="project.asp" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">������Ŀ��Ϣ�ϴ�</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��������������</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">���ʵ����֤�������</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">�����������</span></font></a></li>
					-->
						</td>
    <td width="567" valign="top">
  <table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
  <tr>
        <td nowrap height="500" align="center">
            <table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
                <tr>
                    <td width="1" height="500" align="center" nowrap>
                    	<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
													</table>
                    </td>
					<td align="center" valign="top" height="600" nowrap bgcolor="#99CCFF">
<form name="AddEquip" method="post" action="AddQuote_deal.asp"  enctype="multipart/form-data" onsubmit="return Check()">
<div align="center">
<table cellspacing="0" cellpadding="0" border="1" width="100%" bordercolor="#FFFFFF">
    <tr>
        <td nowrap colspan="2" height="30" align="center" bordercolor="#FFFFFF">
       ��<p><b><font face="��������" size="5">���������Ŀ����</font></b>        </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" width="30%" bordercolor="#FFFFFF">
        ��Ŀ����
        </td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqName">
        </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
        ��Ŀ��Ա
        </td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqProo">
        </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
       ʵ��ʱ�� 
        </td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqAmount">
               </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
       ��Ŀ���
        </td>
	  <td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqPrice">
                </td>
    </tr>
	<tr>
        <td nowrap  height="30" align="center" bordercolor="#FFFFFF">
        ����֤��
        </td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="file" name="EqImage" class="file">(�ϴ��ļ���ʽ��rar)
        </td>
    </tr>
	<tr>
		<td colspan="2" align="center" height="30" bordercolor="#FFFFFF">
		<input type="submit" name="Post" value="�ύ" class="button">
		<input type="reset" name="����" value="����" class="button">
		<input type="button" name="����" value="����" class="button" onclick="Javascript:history.go(-1)">
		</td>
	</tr>
</table>
</div>
</form>	   
   </td>
		  </tr>
</table>
	</td>
  </tr>
  </table>
</div>
  <tr>
    <td width="772" valign="top" colspan="2"> 
        		<!--#include file="bottom.asp"-->	��</td>
    </body>
</html>