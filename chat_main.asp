
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="chat_style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	background-color: #FFEFEF;
}
.style1 {color: #FF0000}
.style2 {
	color: #0000FF;
	font-style: italic;
}
-->
</style>
<script type="text/javascript" src="chat_xmlHttpRequest.js"></script>
<script type="text/javascript">
function scrollWindow(){

                this.scroll(0,75000);
                setTimeout('scrollWindow()',200);
}

</script>
<script type="text/javascript">
timer = window.setInterval("ShowChat()",1000); 
function ShowChat(){
	xmlHttp.open("post","chat_showchat.asp", true);
	xmlHttp.onreadystatechange = function(){
		if(xmlHttp.readyState == 4){
			tet = xmlHttp.responseText;
			document.getElementById("chat_show").innerHTML += tet;
		}
	}
	xmlHttp.send(null);
	//window.open(url);		//测试使用 
}
</script>
</head>

<body onLoad="scrollWindow()">
<font color=#CC0000 size=2>化解困惑，把握心海罗盘；自我调试，拥抱幸福生活。</font><br>
<table>
 <tr>
  <td>
<div id="chat_show"></div>
  </td>
 </tr>
</table>
</body>
</html>
