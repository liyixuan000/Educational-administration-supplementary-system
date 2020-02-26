<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>教务辅助管理信息系统</title>
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
					<font color="#000080">公共社会实践项目</font></span></a> 
					<li><a href="ppro.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">自行组队申请</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">组队立项申请状态查看</span></font></a></li>
					<li><a href="pro.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">社会实践认证提交</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">认证状态查看</span></font></a></li>

					<li><a href="honor.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">社会实践荣誉申请</span></font></a><form method="POST" action="--WEBBOT-SELF--">
						<!--webbot bot="SaveResults" U-File="fpweb:///_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" -->
										</form>
					</li>
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
<form name="AddEquip" method="post" action="honor_deal.asp"  enctype="multipart/form-data" onsubmit="return Check()">
<div align="center">
<table cellspacing="0" cellpadding="0" border="1" width="100%" bordercolor="#FFFFFF">
    <tr>
        <td nowrap colspan="2" height="30" align="center" bordercolor="#FFFFFF">
       　<p><b><font face="华文隶书" size="5">社会实践荣誉申请</font></b>        </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" width="30%" bordercolor="#FFFFFF">
        项目名称
        </td>
		<td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqName">
        </td>
    </tr>
	    
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
     荣誉类型
        </td>
	  <td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="Eqhonor" value="社会实践先进个人" readonly>
                </td>
    </tr>
	<tr>
        <td nowrap  height="30"  align="center" bordercolor="#FFFFFF">
       申请理由
        </td>
	  <td nowrap bordercolor="#FFFFFF">
        <input type="text" class="text" name="EqPrice" size="53">
                </td>
    </tr>
	<tr>
        <td nowrap  height="30" align="center" bordercolor="#FFFFFF">
        材料证明
        </td>
		<td nowrap bordercolor="#FFFFFF">允许文件格式：rar
			        <script type="text/javascript" src="swfupload/swfupload.js"></script>
<script type="text/javascript" src="swfupload/swfupload.handler.js"></script>
<script type="text/javascript">
function StartErrorCallBack(){
	alert("请选择文件!");
	//return false;
	//window.event.returnValue=false;	
}
function StartCallBack(){
	return " 正在上传...";	
}
function ProcessCallBack(bytesComplete, bytesTotal){
	var txt = (bytesComplete/bytesTotal)*100;
	txt = txt.toFixed(2);
	return " [" + txt + "%]";
}
function FailedCallBack(msg){
	return " <img src=\"images/wrong.png\" width=\"16\" height=\"16\" /> [" + msg + "]";	
}
function QueueErrorCallBack(file,msg,message){
	var msg1 = "";
	if(file!=null)msg1=file.name + "\r\n\r\n";
	alert(msg1+msg);
}
function SavingCallBack(){
	return " <img src=\"images/loading.gif\" width=\"16\" height=\"16\" />";	
}
function QueuedCallBack(file){
	return " 等待上传。[<a href=\"javascript:void(0)\" onclick=\"SWFUpload.instances['" + this.movieName + "'].cancelUpload('" + file.id + "');\">取消</a>]";	
}
function CancelledCallBack(id){
	return " 已取消。[<a href=\"javascript:void(0)\" onclick=\"SWFUpload.instances['" + this.movieName + "'].requeueUpload('" + id + "');\">恢复</a>]";
}
function SuccessCallBack(file,serverData){
	var File = null;
	eval("File = (" + serverData + ");");
	if(!File)return"";
	if(File.err){
		swfu.uploadError(file,500,File.msg);
		return "";
	}
	document.getElementById("fileUploaded").innerHTML += "<br />‘" + File.src + "’ 已上传为 ‘" + File.name+"’";
	return " <img src=\"images/right.png\" width=\"16\" height=\"16\" /> [<a href=\"javascript:void(0)\" onclick=\"SWFUpload.instances['" + this.movieName + "'].requeueUpload('" + file.id + "');\">重新上传</a>]";	
}

var Setting={
	debug:false,
	upload_url: "upload.asp",
	file_post_name : "filedata",
	file_types : "*.bmp;*.rar;*.jpg;*.gif;*.png;", //文件格式限制
	file_types_description : "文件类型", //文件格式描述
	file_size_limit : "500 MB", // 文件大小限制
	file_upload_limit : 5, //上传文件总数
	button_append : "divAddFiles",
	button_width: 32,
	button_height: 32,
	button_image_url:"images/btn.png",
	button_text: "<span class=\"but\"></span>",
	button_text_style:".but {color:#ff0000;}",
	button_text_left_padding: 14,
	button_text_top_padding: 3,
	charset: "gbk",
	auto:false
};
var swfu;
window.onload=function(){swfu=HandlerInit(Setting);};
</script>
<span id="divAddFiles"></span>
	
			<!--<input type="image" src="images/cross.png" onclick="swfu.removeAllFiles();" title="取消全部上传队列" width="32" height="32" name="I2" />
			<input type="image" src="images/requeue.png" onclick="swfu.requeueAll();" title="重置全部上传队列" width="32" height="32" name="I3" />-->
		</td>
	</tr>
	<tr>
        <td nowrap  height="30" align="center" bordercolor="#FFFFFF" colspan="2">
		<input type="image" src="images/up.png" onclick="swfu.startUploadFiles();" title="点击开始上传" width="40" height="20" name="I1" align=absbottom />

		<!--<input type="submit" name="Post" value="提交" class="button">-->
		<input type="reset" name="重置" value="重置" class="button">
		<input type="button" name="返回" value="返回" class="button" onclick="Javascript:history.go(-1)">
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
        		<!--#include file="bottom.asp"-->	　</td>
    </body>
</html>