<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="chat_style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	background-color: #FFECEC;
}
-->
</style>
<script type="text/javascript" src="chat_xmlHttpRequest.js"></script>
<script type="text/javascript">
function time(){
	tim = new Date();
	var strDate = tim.getYear() + "-";
	strDate += tim.getMonth() + "-";
	strDate += tim.getDay() + " ";
	strDate += tim.getHours() + ":";
	strDate += tim.getMinutes() + ":";
	strDate += tim.getSeconds();
	return strDate;
}
function talk(){
	 mess = form1.message.value;
	if(mess==""){
		alert("�������ݲ���Ϊ�գ�");
		form1.message.focus();
		return false;
	}
	else{
	//my = document.getElementById("from").value;
	//to = document.getElementById("to").value;
	//face = document.getElementById("face").value;
	color = document.getElementById("select").value;
	message = document.getElementById("message").value;
	content +=" <i><font color='#0000FF'>";
	//content += "</font></i> ��[<font color='#FF0000'>"+to+"</font>] ˵��";
	content += "<font color=#"+color+">"+message+"</font> ("+time()+")";
	document.getElementById("message").value = "";
	url = "chat_shownow.asp?content=" + content;
	xmlHttp.open("get",url, true);
	xmlHttp.send();}
}
</script>
</head>
<body on>
<form method="post" action="chat_main.asp" target="mainFrame" name="form1">
      <input name="from" type="hidden" id="from" value="<%=session("UserName")%>">
  <table width="100%" height="70"  border="1" cellpadding="0" cellspacing="0" bgcolor="#FFDFDF" bordercolordark="#000000" bordercolorlight="#FFFFFF" bordercolor="#000000">
    <script language="javascript">
	function Exit(){
	parent.location.href="chat_exit.asp";
	}
	</script>
    <tr>
      <td height="35">&nbsp;
	  	<input name="message" type="text" id="message" size="60" onKeyDown="Javascript:if(event.keyCode==13){talk()}">
      	<input name="send" type="button" id="send" value="����" onClick="talk();return false">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		������ɫ��
          <select name="color" size="1" id="select">
            <option selected value="000000">Ĭ����ɫ</option>
            <option style="color:#FF0000" value="FF0000">��ɫ����</option>
            <option style="color:#0000FF" value="0000ff">��ɫ����</option>
            <option style="color:#ff00ff" value="ff00ff">��ɫ����</option>
            <option style="color:#009900" value="009900">��ɫ�ഺ</option>
            <option style="color:#009999" value="009999">��ɫ��ˬ</option>
            <option style="color:#990099" value="990099">��ɫ�н�</option>
            <option style="color:#990000" value="990000">��ҹ�˷�</option>
            <option style="color:#000099" value="000099">��������</option>
            <option style="color:#999900" value="999900">�����Ʒ�</option>
            <option style="color:#ff9900" value="ff9900">�ֽ�����</option>
            <option style="color:#0099ff" value="0099ff">��������</option>
            <option style="color:#9900ff" value="9900ff">��������</option>
            <option style="color:#ff0099" value="ff0099">���İ�ʾ</option>
            <option style="color:#006600" value="006600">ī�����</option>
            <option style="color:#999999" value="999999">��������</option>
        </select></td>
      <td><div align="center">
          <input type="button" name="Submit2" value="�˳�������ѯ��" onClick="Exit()">
      </div></td>
    </tr>
  </table>
</form>
</body>
</html>
