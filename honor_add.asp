<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
dim pclass_id:pclass_id=checkint(request("pclass_id"),0)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>������������Ϣϵͳ</title>
<link rel="stylesheet" href="css/css.css"  type="text/css"><script type="text/javascript" src="js/laydate.js"></script>

<style type="text/css">
*{margin:0;padding:0;list-style:none;}
html{background-color:#E3E3E3; font-size:14px; color:#000; font-family:'΢���ź�'}
h2{line-height:30px; font-size:20px;}
a,a:hover{ text-decoration:none;}
pre{font-family:'΢���ź�'}
.box{width:970px; padding:10px 20px; background-color:#fff; margin:10px auto;}
.box a{padding-right:20px;}
.demo1,.demo2,.demo3,.demo4,.demo5,.demo6{margin:25px 0;}
h3{margin:10px 0;}
.layinput{height: 22px;line-height: 22px;width: 150px;margin: 0;}
</style>
</head>
<body><script src="area_Ajax.js" ></script>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="35" bgcolor="#ECF5FF">&nbsp;��ӭ����<font color="#FF0000"><%=session("loginusername")%></font>,�����ڵ������<font color="#FF0000"><%if session("power")=1 then
	response.Write("��������Ա")
	elseif session("power")=2 then
	response.Write("��ͨ����Ա")
	else
	response.Write("ѧ��")
	end if%></font></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td width="120" valign="top"><table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="practice_manage.asp" target="main"> <span > <font color="#fff">�������ʵ����Ŀ</font></span></a></td>
      </tr>
      <%if session("power")=3 then%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="team_add.asp" target="main"><font color="#fff"> <span >�����������</span></font></a>
          </li></td>
      </tr>
      <%end if%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="team_manage.asp" target="main"> <font color="#fff"> <span >�����������״̬�鿴</span></font></a></td>
      </tr>
      <%if session("power")=3 then%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="certified_add.asp" target="main"><font color="#fff"> <span >���ʵ����֤�ύ</span></font></a>
          </li></td>
      </tr>
      <%end if%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="certified_manage.asp" target="main"> <font color="#fff"> <span >��֤״̬�鿴</span></font></a></td>
      </tr>
      <%if session("power")=3 then%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="honor_add.asp" target="main"><font color="#fff"> <span >���ʵ����������</span></font></a>
          </li></td>
      </tr>
      <%end if%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="honor_manage.asp" target="main"> <font color="#fff"> <span >��������״̬�鿴</span></font></a></td>
      </tr>
      <%if session("power")=2 then%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="practice_add.asp" target="main"><font color="#fff"> <span >�������ʵ����Ŀ�ϴ�</span></font></a></td>
      </tr>
      <%end if%>
    </table></td>
    <td width="567" valign="top"><table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
      <tr>
        <td nowrap="nowrap" height="500" align="center"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
          <tr>
            <td width="1" height="500" align="center" nowrap="nowrap"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
            </table></td>
            <td align="center" valign="top" height="600" nowrap="nowrap" bgcolor="#99CCFF"><form action="honor_save.asp?action=Add" method="post" name="myform"onsubmit="return check()">
              <div align="center">
                <table cellspacing="0" cellpadding="0" border="1" width="100%" bordercolor="#FFFFFF">
                  <tr>
                    <td nowrap="nowrap" colspan="2" height="30" align="center" bordercolor="#FFFFFF"> ��
                      <p><b><font size="5" face="��������"><strong>���ʵ������</strong>����</font></b></p></td>
                  </tr>
                  <tr>
                    <td width="30%"  height="30"  align="right" nowrap="nowrap" bordercolor="#FFFFFF">��Ŀ���ƣ� </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input type="text" class="text" name="bnumber" />
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr> <tr>
                    <td nowrap="nowrap"  height="30"  align="right" bordercolor="#FFFFFF">�������ͣ�</td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><input name="title" type="text" class="text" id="title" />
                      <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                  </tr>
                           <tr>
                             <td nowrap="nowrap"  height="30"  align="right" bordercolor="#FFFFFF">�������ɣ� </td>
                             <td nowrap="nowrap" bordercolor="#FFFFFF"><textarea name="content" rows="6" class="text" id="content"></textarea>
                             <span class="category"> &nbsp;<font color="#ff0000">*</font></span></td>
                          </tr><tr>
    <td height="30" align="right">����֤����</td>
    <td class="category">
        <div class="style2" id=picUrl_display style="display:none;"><input name="picUrl" type="text" id="picUrl" size="50"></div>
				  <iframe src="picUpload.asp" width=100% height=40 frameborder="0" scrolling="no"></iframe></td>
  </tr>
                  <tr>
                    <td colspan="2" align="center" height="30" bordercolor="#FFFFFF"><input type="submit" name="Post" value="�ύ" class="button" />
                      <input type="reset" name="����" value="����" class="button" />
                      <input type="button" name="����" value="����" class="button" onClick="Javascript:history.go(-1)" /></td>
                  </tr>
                </table>
              </div>
            </form></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="778" colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</div>
</body>
</html>

<script>
function check()
{
	if(myform.bnumber.value=='')
	{
		alert("��Ŀ���Ʋ���Ϊ�գ�");
		myform.bnumber.focus();
		return false;
	}
	if(myform.title.value=='')
	{
		alert("�������Ͳ���Ϊ�գ�");
		myform.title.focus();
		return false;
	}
	if(myform.content.value=='')
	{
		alert("�������ɲ���Ϊ�գ�");
		myform.content.focus();
		return false;
	}
}
</script>