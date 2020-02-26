<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Dim conn
Set conn=Server.CreateObject("ADODB.connection")
sql="Driver={Microsoft Access Driver (*.mdb)};DBQ= "&server.MapPath("database/db_database.mdb")
conn.open(sql)
%>
<% 
	dim rndnum,verifycode
	Randomize
	Do While Len(rndnum)<4
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
	loop
	session("verifycode")=rndnum
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>教务辅助管理信息系统用户登录</title>
<script language="javascript">
function Mycheck(){
	if (myform.username.value=='')
	{
		alert("用户名不能为空！！");
		myform.username.focus();
		return false;
	}
	if(myform.password.value=="")
	{
		alert("用户密码不能为空！！");
		myform.password.focus();
		return false;
	}
	if(myform.yan.value=="")
	{
		alert("请输入随机产生的验证码！！");
		myform.yan.focus();
		return false;
	}
}
</script>
<style type="text/css">
<!--
.style2 {
	font-size: 9pt;
	color: #000000;
}
.style5 {font-size: 9pt}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_setTextOfTextfield(objName,x,newText) { //v3.0
  var obj = MM_findObj(objName); if (obj) obj.value = newText;
}
//-->
</script>
</head>
<body>
<%
	if request("action")="check" then
	username=request.form("username")
	password=request.form("password")
	for i=1 to len(username)
	username2=mid(username,i,1)
	if username2="'" or username2="%" or username2="<" or username2=">" or username2="&" or username2="|" then 
	response.write "<script language=jscript>alert('您的用户名含有非法字符,请重新输入！');history.back()</script>"
	response.end
	end if 
	next
	for i=1 to len(password)
	password2=mid(password,i,1)
	if  password2="'" or password2="%" or password2="<" or password2=">" or password2="&" or password2="|" then
	response.Write"<script language=jscript>alert(您的密码含有非法字符,请重新输入！);history.back()</script>"
	response.End()
	end if 
    next
%>
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select top 1 * from userinfo where username='"&username&"' and password='"&password&"'"
'response.write sql
rs.open sql,conn,1,3
if rs.eof then
response.write "<script language=jscript>alert('对不起，您输入的用户名、密码或验证码有误，请重新输入，谢谢！');window.location.href='index.asp';</script>"
else
username=request.form("username")
session("loginusername")=rs("username")
session("power")=rs("power")
response.write "<script>alert('登录成功！！');window.location.href='home.asp';</script>"
end if
rs.close
set rs=nothing
end if
%>
 <p>　</p>
  <p>　</p>
   <p>　</p>
<form name="myform" method="post" action="?action=check" onSubmit="return Mycheck();">
<table width="402" height="137" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#FFCCFF" bordercolordark="#99FFFF">
  <tr>
    <td width="398" valign="top" background="images/1.gif"><table width="391" height="137" border="0" align="right" cellpadding="0" cellspacing="0">
      <tr>
        <td width="214" height="76">　</td>
        <td width="49" valign="bottom"><span class="style2">用户名：</span>          </td>
        <td width="128" valign="bottom"><input name="username" type="text" id="username" size="16" onFocus="this.select(); "style="border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1;width:120px;height:15px" onmouseover="this.style.background='#ccffff';" onmouseout="this.style.background='#FFFFFF'"></td>
      </tr>
      <tr>
        <td height="30">　</td>
        <td height="19" class="style2">密　码：          </td>
        <td height="19" class="style2"><input name="password" type="password" id="password" size="16" onFocus="this.select(); "style="border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1;width:120px;height:15px" onmouseover="this.style.background='#ccffff';" onmouseout="this.style.background='#FFFFFF'"></td>
      </tr>
      <tr>
        <td height="11">&nbsp;</td>
        <td height="11" class="style2">验证码：          </td>
        <td height="11" class="style2"><input name="yan" type="text" id="yan" size="5" onFocus="this.select(); " style="border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1;width:50px;height:15px"onmouseover="this.style.background='#ccffff';" onmouseout="this.style.background='#FFFFFF'">
          <span class="style5"><font color="#000000"><%=session("verifycode")%>
          <input name="hcheckCode" type="hidden" id="hcheckCode" value="<%=checkCode%>">
          </font></span></td>
      </tr>
      <tr>
        <td height="27">　</td>
        <td height="27" colspan="2" class="style2"><div align="center"><input name="submit" type="image" src="images/2.gif" >　<a href="javascript:;" onClick="MM_setTextOfTextfield('name1','','');MM_setTextOfTextfield('pwd','','')"><img src="images/3.gif" alt="" border="0"></a></div></td>
        </tr>
    </table>
    <p>　</p>
    <p>　</p></td>
  </tr>
</table>
</form>
</body>
</html>
