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
	background-color: #FFECEC;
}
-->
</style>
<script type="text/javascript" src="chat_xmlHttpRequest.js"></script>
<script type="text/javascript">
timer = window.setInterval("ShowAll()",2000);
function ShowAll(){
xmlHttp.open("post","chat_show.asp", true);
	xmlHttp.onreadystatechange = function(){
		if(xmlHttp.readyState == 4){
			tel = xmlHttp.responseText;
			document.getElementById("show").innerHTML = tel;
		}
	}
	xmlHttp.send(); 
	//window.open(url);		//测试使用 
}
</script>
</head>
<body>
<p>
欢迎您走进心理咨询室！<br>
<div id="show"></div>

</body>
</html>
