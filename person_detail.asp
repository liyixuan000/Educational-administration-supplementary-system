<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="user_check.asp"-->
<%
Dim Id:Id=checkint(request("id"),0)
dim pclass_id:pclass_id=checkint(request("pclass_id"),0)
if Id=0 then jstop("��ܰ��ʾ����������")
dim rsEdit
dim rs,sqlstr,Class_id,picurl,postdate,content,title,pagetitle,pclass_id2,bnumber,orders,prize,pid,username,addtimes
sqlstr = "select * from person where id="&Id
set rsEdit=conn.execute(sqlstr)
bnumber =rsEdit("bnumber")
picurl=rsEdit("picurl")
pid=rsEdit("pid")
orders=rsEdit("orders")
prize=rsEdit("prize")
title=rsEdit("title")
content=rsEdit("content")
Class_id=rsEdit("Class_id")
username=rsEdit("username")
addtimes=rsEdit("addtimes")
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
    <td height="35" bgcolor="#ECF5FF">&nbsp;��ӭ����<font color="#FF0000"><%=session("loginusername")%></font>,�����ڵ�������<font color="#FF0000"><%if session("power")=1 then
	response.Write("��������Ա")
	elseif session("power")=2 then
	response.Write("��ͨ����Ա")
	else
	response.Write("ѧ��")
	end if%></font></td>
  </tr>
</table><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100" height="35" bgcolor="#FFFFFF">&nbsp;<strong>���Ϸ��ࣺ<font color="#FF0000"></font></strong></td>
    <td bgcolor="#FFFFFF"><table  border="0" cellspacing="1" cellpadding="1">
      <tr>        <%
				dim rsc,selected
			 	set rsc=conn.execute ("select class_id,class_name_"&lan_txt&",class_depth from Article002class where parent_id=0  order by class_order desc,class_id asc")
				do while not rsc.eof

			 %>
        <td width="80" height="35" align="center" <%if rsc("class_id")=class_id then 
		response.Write("bgcolor='#1496F5'")
		else
		response.Write("bgcolor='#C8EDFB'")
		end if
		%>
		><a href="data_manage.asp?class_id=<%=rsc("class_id")%>&id=<%=id%>"><%=rsc("class_name_"&lan_txt&"")%></a></td>
        <%

					rsc.movenext
				loop
				rsc.close:set rsc=nothing
				%><td width="80" height="35" align="center" bgcolor='#C8EDFB'><a href="person_manage.asp">������Ϣ</a></td>

      </tr>
    </table></td>
  </tr>
</table></td>
  </tr>
  <tr>
    <td width="120" valign="top"><table width="98%" border="0" align="center" cellpadding="1" cellspacing="1">
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="data_manage.asp" target="main"> <span > <font color="#fff">�����б�</font></span></a></td>
      </tr>
      <%if session("power")=3 then%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="person_add.asp" target="main"><font color="#fff"> <span >����������Ϣ</span></font></a>
          </li></td>
      </tr>
      <%end if%>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="person_manage.asp" target="main"> <font color="#fff"> <span >������Ϣ�����б�</span></font></a></td>
      </tr>      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="data_add.asp" target="main"><font color="#fff"> <span >��������</span></font></a></td>
      </tr>
      <%if session("power")=2 then%>

      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="dataclass_add.asp" target="main"><font color="#fff"> <span >�������Ϸ���</span></font></a></td>
      </tr>
      <tr>
        <td height="25" bgcolor="#39B2DD"><a href="dataclass_manage.asp" target="main"><font color="#fff"> <span >�������Ϸ���</span></font></a></td>
      </tr>
      <%end if%>
    </table></td>
    <td width="567" valign="top"><table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">
      <tr>
        <td nowrap="nowrap" height="500" align="center"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
          <tr>
            <td width="1" height="500" align="center" nowrap="nowrap"><table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%" align="center">
            </table></td>
            <td align="center" valign="top" height="600" nowrap="nowrap" bgcolor="#99CCFF"><form action="person_save.asp?action=Edit" method="post" name="myform"onsubmit="return check()">
              <div align="center">
                <table cellspacing="0" cellpadding="0" border="1" width="100%" bordercolor="#FFFFFF">
                  <tr>
                    <td nowrap="nowrap" colspan="2" height="30" align="center" bordercolor="#FFFFFF"> ��
                      <p><b><font size="5" face="��������">������Ϣ��������</font></b></p></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="right" width="30%" bordercolor="#FFFFFF"> ������� </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><span class="category">
                      <%
	  dim sqlC,class_name,parent_id
	  sqlC="select class_name_"&lan_txt&",parent_id from Article002Class where  class_id="&class_id&""
	  set rsc=conn.execute (sqlC)
	  if not rsC.eof then
	  	class_name=rsc("class_name_"&lan_txt&"")
	  	response.write class_name
	  end if
	  %></span></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="right" bordercolor="#FFFFFF"> ѧ�ţ� </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><%=bnumber%></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="right" bordercolor="#FFFFFF"> ��ϵ��ʽ�� </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><%=prize%></td>
                  </tr>
                  <tr>
                    <td nowrap="nowrap"  height="30"  align="right" bordercolor="#FFFFFF"> �������ƣ� </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><%=title%></td>
                  </tr>                  <tr>
                    <td nowrap="nowrap"  height="30"  align="right" bordercolor="#FFFFFF"> ������Ϣ��飺 </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><%=content%></td>
                  </tr>   <tr>
                    <td nowrap="nowrap"  height="30"  align="right" bordercolor="#FFFFFF"> �����ˣ� </td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><%=username%></td>
                  </tr>
<tr>
                    <td nowrap="nowrap"  height="30" align="right" bordercolor="#FFFFFF">����ʱ�䣺</td>
                    <td nowrap="nowrap" bordercolor="#FFFFFF"><%=infotime(addtimes)%></td>
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
	if(myform.class_id.value=='0')
	{
		alert("���������Ϊ�գ�");
		myform.class_id.focus();
		return false;
	}
	if(myform.bnumber.value=='')
	{
		alert("ѧ�Ų���Ϊ�գ�");
		myform.bnumber.focus();
		return false;
	}
	if(myform.prize.value=='')
	{
		alert("��ϵ��ʽ����Ϊ�գ�");
		myform.prize.focus();
		return false;
	}
	if(myform.title.value=='')
	{
		alert("�������Ʋ���Ϊ�գ�");
		myform.title.focus();
		return false;
	}
	if(myform.content.value=='')
	{
		alert("������Ϣ��鲻��Ϊ�գ�");
		myform.content.focus();
		return false;
	}
}
</script>