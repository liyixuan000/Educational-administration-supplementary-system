<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
session.Timeout=10    '����Session����������
flag=false
'�û���¼
If Request.Form("Submit")="����������" Then
	UserName=trim(Request.Form("UserName"))
	If UserName="" Then
	response.Write "<script>alert('�������ǳ�');location='chat_index.asp';</script>"
	End If
	'�û�ͷ��
	head=Request.Form("head")
	'����û���
	A_Username=split(Application("UserName"),",")
	for i=0 to uBound(A_Username)
		If A_username(i)=username then
			response.Write "<script>alert('�ǳ��ظ�');location='chat_index.asp';</script>"
		flag=true
		End if
	Next
	'�洢�û���Ϣ������������
	If flag=false Then
		application.Lock()
		application("UserName")=application("UserName")+","+UserName
		session("UserName")=UserName
		Application("head")=application("head")+","+head
		application("newmen") = application("newmen") + 1
		application.UnLock()
		response.Redirect("chat_chatroom.asp")
	End if
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>������ѯ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="chat_style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	background-color: #FFEFEF;
	background-image: url(bg.jpg);
}
.style1 {color: #990000}
-->
</style></head>

<body>
<form method="post" action="chat_index.asp" name="form1">
  <table width="650" height="307"  border="0">
    <tr>
      <td width="335" align="center">��</td>
    </tr>
    <tr>
      <td valign="top" align="center"><table width="332" height="163" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="38" height="100">��</td>
          <td width="294" valign="top">		  <table width="271" height="100"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFDFDF">
            <tr>
              <td width="33" height="33" align="center">                  <input name="head" type="radio" value="1.GIF" checked="checked" />                </td>
              <td width="57" height="33" align="center"><div align="center"> &nbsp;<img src="chat_headICO/1.GIF" width="32" height="32" /></div></td>
              <td width="33" height="33" align="center"><input name="head" type="radio" value="2.GIF" /></td>
              <td width="57" height="33"><div align="center"><img src="chat_headICO/2.GIF" width="32" height="32" /></div></td>
              <td width="33" align="center"><input name="head" type="radio" value="3.GIF" /></td>
              <td width="57"><div align="center"><img src="chat_headICO/3.GIF" width="32" height="32" /></div></td>
              </tr>
            <tr>
              <td width="33" height="33" align="center">                  <input name="head" type="radio" value="5.GIF" />                </td>
              <td width="57" height="33" align="center"><div align="center"><img src="chat_headICO/5.GIF" width="37" height="34" /></div></td>
              <td width="33" height="33" align="center"><input name="head" type="radio" value="4.GIF" /></td>
              <td width="57" height="33"><div align="center"><img src="chat_headICO/4.GIF" width="33" height="32" /></div></td>
              <td width="33" align="center"><input name="head" type="radio" value="9.GIF" /></td>
              <td width="57"><div align="center"><img src="chat_headICO/9.GIF" width="32" height="32" /></div></td>
              </tr>
            <tr>
              <td width="33" height="33" align="center">                  <input name="head" type="radio" value="7.GIF" />                </td>
              <td width="57" height="33" align="center"><div align="center"><img src="chat_headICO/7.GIF" width="35" height="33" /></div></td>
              <td width="33" height="33" align="center"><input name="head" type="radio" value="6.GIF" /></td>
              <td width="57" height="33"><div align="center"><img src="chat_headICO/6.GIF" width="32" height="32" /></div></td>
              <td width="33" align="center"><input name="head" type="radio" value="8.GIF" /></td>
              <td width="57"><div align="center"><img src="chat_headICO/8.GIF" width="35" height="34" /></div></td>
              </tr>
          </table></td>
        </tr>
	<script language="javascript">
	function check(){
		if(form1.Username.value==""){
			alert("�ǳƲ���Ϊ�գ�");
			form1.Username.focus();return false;
		}
	}
	</script>
        <tr>
          <td height="33">��</td>
          <td height="27"><div align="center"><span class="style1">�ǳƣ�</span>                
            <input name="Username" type="text" class="Sytle_text" id="Username3" onKeyDown="Javascript:if(event.keyCode==13){ return check()}" />
          </div></td>
        </tr>
        <tr>
          <td height="27">��</td>
          <td height="27"><div align="center">
              <input name="Submit" type="Submit" class="Style_button_del" value="����������" onclick="return check()"/>
          </div></td>
        </tr>
      </table>        </td>
    </tr>
    <tr>
      <td align="center">��</td>
    </tr>
  </table>
</form>
</body>
</html>
