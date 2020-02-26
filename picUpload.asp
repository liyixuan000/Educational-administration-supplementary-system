<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<!--#include file="inc/upload_5xsoft.asp"-->
<%

dim upload,files,sexName,exName,years,months,days,hours,minutes,seconds,Dir,path,picUrl_style
set upload = new upload_5xsoft
set files = upload.file("picUrl")
if files.filesize>0 then
	sexName=split(files.filename,".")
	exName="."&sexName(ubound(sexName))
	exName=lcase(exName)
	'if exName<>".gif" or exName<>".jpg" or exName<>".bmp" or exName<>".jpeg"  then
'		response.write "图片格式不对！<a href="&request.ServerVariables("http_referer")&"><font color=blue>重新上传</font></a>"
'		response.End()
'	end if
	years = year(now)
    months = right("00"&month(now),2)
    days = right("00"&day(now),2)
    hours = right("00"&hour(now),2)
    minutes = right("00"&minute(now),2)
    seconds = right("00"&second(now),2)
    Dir = years&months&days&hours&minutes&seconds&exName
	path="/files/"&Dir
	'response.Write(path)
	files.saveAs server.mappath(path)
	response.write "<script>parent.document.all.picUrl_display.style.display='block';parent.document.all.picUrl.value='"&path&"';</script>"
	picUrl_style="none"
end if
set files = nothing
set upload = nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style type="text/css">
<!--
td,body{font-size:9pt;}
body {
	background-color: #99CCFF;
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script language="javascript">
<!--
function checkform(form)
{
	if(form.picUrl.value.replace(/\s/gi,"") ==""){alert("请选择上传的文件");return false;}
	if(form.picUrl.value.replace(/\s/gi,"") !="")
	{
	    //var fso = new ActiveXObject("Scripting.FileSystemObject");
		//fileName = fso.FileExists(form.keJian_picUrl.value);
		//if(! fileName){alert("请选择上传的文件");form.keJian_picUrl.value="";return false;} 
		var sstr = form.picUrl.value.split(".")
		var str ="."+sstr[sstr.length-1].toLowerCase();
		if (str != ".doc" && str != ".rar" && str != ".zip" && str != ".pdf" && str != ".docx"){alert("只允许上传的文件格式为:doc|rar|zip|pdf");return false;}
	}
	
	return true;
}
//-->
</SCRIPT>
</head>

<body >
<div id="picUrl" style="display:<%=picUrl_style%>;">
<form action="" method="post" enctype="multipart/form-data" name="form1" onSubmit="return checkform(this)">
<input name="picUrl" type="file" id="keJian_picUrl2" size="40">
  <input name="Submit22" type="submit" id="Submit222" value="上传">（支持doc|rar|zip|pdf四种类型的文件）
</form></div>
<%
if picUrl_style<>"" then response.write "文件上传成功！<a href="&request.ServerVariables("http_referer")&"><font color=blue>重新上传</font></a>"
%>
</body>
</html>
