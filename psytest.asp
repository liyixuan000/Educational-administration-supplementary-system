<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="include/config1.asp"-->
<!--#include file="include/conn1.asp"-->

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

                    <<li><a href="#" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">���֤����</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��ӻ</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">��ӻ��</span></font></a></li>
					
					</td>-->

    <td width="778" valign="top">
<table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
<!--��ֻ��һ�����-->
</script>
<table cellSpacing=0 cellPadding=0 align=center background="images/topbj.jpg" width="778">
  <!--<tr>
    <td colSpan=3 id="banner"><%=SiteName%></td>-->
  </tr>
  <tr>
    <td colSpan=3 id="menu" bgcolor="#0099FF">����֢�Բ���</td>
  </tr>
</table>
<div align="center">
<table width="100%" bgcolor="#FAFAFA" border="1" bordercolor="#FFFFFF">
<form id="form1" name="fm" method="post" action="#" onsubmit="return checkForm(this)"> 
  <tr>
   <td colspan="4" id="BigTitle" bgcolor="#99CCFF">һ��ѡ���� <%=ChoiceDes%> </td>
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
      <td width="50%"><input type="radio" name="Choice<%=k%>" value="A" title="��ѡ��ѡ����<%=k%>�⡱~!" />
        A��<%=rsChoice("A")%></td>
      <td><input type="radio" name="Choice<%=k%>" value="B" title="��ѡ��ѡ����<%=k%>�⡱~!" />
        B��<%=rsChoice("B")%></td>
    </tr>
    <tr>
      <td width="50%"><input type="radio" name="Choice<%=k%>" value="C" title="��ѡ��ѡ����<%=k%>�⡱~!" />
        C��<%=rsChoice("C")%></td>
      <td><input type="radio" name="Choice<%=k%>" value="D" title="��ѡ��ѡ����<%=k%>�⡱~!" />
        D��<%=rsChoice("D")%></td>
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
   <td colspan="4" id="BigTitle" bgcolor="#99CCFF"><br>������ѡ�� <%=MChoiceDes%> </td>
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
      <td width="50%"><input type="checkbox" name="MChoice<%=m%>" value="A" title="��ѡ���ѡ����<%=m%>�⡱~~min:1" />
        A��<%=rsMChoice("A")%></td>
      <td><input type="checkbox" name="MChoice<%=m%>" value="B" title="��ѡ���ѡ����<%=m%>�⡱~~min:1" />
        B��<%=rsMChoice("B")%></td>
    </tr>
    <tr>
      <td width="50%"><input type="checkbox" name="MChoice<%=m%>" value="C" title="��ѡ���ѡ����<%=m%>�⡱~~min:1" />
        C��<%=rsMChoice("C")%></td>
      <td><input type="checkbox" name="MChoice<%=m%>" value="D" title="��ѡ���ѡ����<%=m%>�⡱~~min:1" />
        D��<%=rsMChoice("D")%></td>
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
   <td colspan="4" id="BigTitle" bgcolor="#99CCFF"><br>�����ж��� <%=JudgeDes%> </td>
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
	  <input type="radio" name="Judge<%=w%>" value="A" title="��ѡ���жϡ���<%=w%>�⡱~!" />�� 
	  <input type="radio" name="Judge<%=w%>" value="B" title="��ѡ���жϡ���<%=w%>�⡱~!" />��</td>
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
	<input type="submit" name="Submit" value="�ύ���Խ��" onclick="return checkClass();" style="width:115; height:42; font-size:18px"/></td>
  </tr>
</form>
</table>
</div>
<%
conn.close
set conn=nothing
%>
<!--��Ҳ��һ�����-->
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