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
		alert("发言内容不能为空！");
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
	//content += "</font></i> 对[<font color='#FF0000'>"+to+"</font>] 说：";
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
      	<input name="send" type="button" id="send" value="发送" onClick="talk();return false">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		字体颜色：
          <select name="color" size="1" id="select">
            <option selected value="000000">默认颜色</option>
            <option style="color:#FF0000" value="FF0000">红色热情</option>
            <option style="color:#0000FF" value="0000ff">蓝色开朗</option>
            <option style="color:#ff00ff" value="ff00ff">桃色浪漫</option>
            <option style="color:#009900" value="009900">绿色青春</option>
            <option style="color:#009999" value="009999">青色清爽</option>
            <option style="color:#990099" value="990099">紫色拘谨</option>
            <option style="color:#990000" value="990000">暗夜兴奋</option>
            <option style="color:#000099" value="000099">深蓝忧郁</option>
            <option style="color:#999900" value="999900">卡其制服</option>
            <option style="color:#ff9900" value="ff9900">镏金岁月</option>
            <option style="color:#0099ff" value="0099ff">湖波荡漾</option>
            <option style="color:#9900ff" value="9900ff">发亮蓝紫</option>
            <option style="color:#ff0099" value="ff0099">爱的暗示</option>
            <option style="color:#006600" value="006600">墨绿深沉</option>
            <option style="color:#999999" value="999999">烟雨蒙蒙</option>
        </select></td>
      <td><div align="center">
          <input type="button" name="Submit2" value="退出心理咨询室" onClick="Exit()">
      </div></td>
    </tr>
  </table>
</form>
</body>
</html>
