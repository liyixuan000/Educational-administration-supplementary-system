<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="inco/conn.asp"-->
<!--#include file="inco/config.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<link rel="stylesheet" href="css/css.css"  type="text/css">
<script language=JavaScript>
<!--
function check()
{

  if (document.add.xx1.value=="")
     {
      alert("����дԤ��ʱ�䣡")
      document.add.xx1.focus()
      document.add.xx1.select()
      return
     }
	 
  if (document.add.xx2.value=="")
     {
      alert("����д���ڣ�")
      document.add.xx2.focus()
      document.add.xx2.select()
      return
     }
	 
  if (document.add.xx3.value=="")
     {
      alert("����дԤ����ʱ��")
      document.add.xx3.focus()
      document.add.xx3.select()
      return
     }
	 
  
 	 
  if (document.add.xx5.value=="")
     {
      alert("����д��ϵ�绰��")
      document.add.xx5.focus()
      document.add.xx5.select()
      return
     }
	 	 
     document.add.submit()
}
-->
</script>

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
 </table>
<table width="778" align="center" bgcolor="#F0F7FF">
<%if ks=1 then %>
 <tr>
  <td id="info" colspan="4">����д������ԤԼ��Ϣ��Ȼ��㡰�ύԤԼ����Ť��</td>
 </tr>
 <tr>
  <td id="info2" colspan="4"></td>
 </tr>
 <tr>
  <td colspan="4" style="line-height:35px">
   <table width="100%" align="center" bgcolor="#FAFAFA"> 
     <form name="add" method="post" action="submit.asp">
	  <tr bgcolor="#FFFFFF">
		  <td width="20%" align="right">ԤԼ���ڣ�</td>
		  <td width="79%"><input name="xx1" type="text" id="xx1" value="<%=request("rq")%>" readonly> </td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>�ǡ����ڣ�</td>
		  <td width="79%"><input name="xx2" type="text" id="xx2" value="<%=request("xq")%>" readonly/></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>ԤԼʱ�䣺</td>
		  <%
		  %>
		  <td width="79%"><input name="xx3" type="text" id="xx3"  value="��<%=request("ks")%>��" readonly/></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>ԤԼҽʦ��</td>
          <td width="79%">
		  <SELECT name='xx4' id="xx4">
		  <option value="<%=request("cid")%>"><%=request("cid")%></option>
		  </SELECT></td>

		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>��ϵ�绰��</td>
		  <td width="79%"><input name="xx5" type="text" id="xx5"></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>��ע˵����</td>
		  <td width="79%"  ><textarea name="xx6" cols="53" rows="3" id="xx6"></textarea></td>
		</tr>
	    <tr bgcolor='#FFFFFF'>
		  <td align='right' >����ʱ�䣺</td>
		  <td width="79%" ><input name="addtime" type="text" id="addtime" value="<%=now()%>" readonly></td>
		 </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td colspan="2" ><input TYPE="hidden" name="action" value="yes"><input type="button" name="Submit" value="�ύԤԼ" onClick="check()"></td>
        </tr>
		</form>
      </table>
  </td>
 </tr>
<%end if%>
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