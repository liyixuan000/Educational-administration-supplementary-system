<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
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
    <td colspan="2"><!--#include file="top2.asp"--></td>
  </tr>
  <tr>
    <td width="205" valign="top"> 
        			<li><a href="data.asp" target="main">
					<span style="background-color: #C0C0C0">
					<font color="#000080">��������</font></span></a> 
					<li><a href="activitycj_add.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">�ѷ���������Ϣ</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">������Ϣ�ϴ�</span></font></a></li>
					<li><a href="per_add.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">����������Ϣ</span></font></a></li>
                    <!--<li><a href="#" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">���֤����</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��ӻ</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��ӻ��</span></font></a></li>
					-->
					</td>
    <td width="567" valign="top">
 </div>
</form>	 
 
<%namee=request("EqName")
'response.write(name)
Phone=request("EqPhone")
Name=request("EqAmount")
EClasss=request("Eqppro")
intro=request("EqPrice")
%>  


<%	conn = Server.CreateObject("ADODB.Connection")
sql="Driver={Microsoft Access Driver (*.mdb)};DBQ= "&server.mappath("Database/db_database.mdb")
	rs = Server.CreateObject("ADODB.Recordset")	
	sql="insert into tb_per (namee,Name1,Phone,Intro,EClass) values('"+namee+"','"+Name+"','"+Phone+"','"+Intro+"','"+EClass+"')"

	//sql="insert into goumai set shuliang1 = shuliang1 +"+shuliang1+",price1 =price1+"+price1+",shuliang2 =shuliang2+"+shuliang2+",price2 =price2+"+price2+" ,shuliang3 =shuliang3+"+shuliang3+",price3 =price3+"+price3+",shuliang4 =shuliang4+"+shuliang4+",price4 =price4+"+price4+",shuliang5 =shuliang5+"+shuliang5+",price5 =price5+"+price5+" where xingming='"+xingming+"'"
	Response.Write(sql);
	//conn.Execute(sql)

%>

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