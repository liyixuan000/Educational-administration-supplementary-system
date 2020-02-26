<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!--#include file="inco/conn.asp"-->
<!--#include file="inco/config.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<link rel="stylesheet" href="css/css.css"  type="text/css">
<script language=JavaScript>
<!--
function check()
{

  if (document.add.xx1.value=="")
     {
      alert("请填写预定时间！")
      document.add.xx1.focus()
      document.add.xx1.select()
      return
     }
	 
  if (document.add.xx2.value=="")
     {
      alert("请填写星期！")
      document.add.xx2.focus()
      document.add.xx2.select()
      return
     }
	 
  if (document.add.xx3.value=="")
     {
      alert("请填写预定课时！")
      document.add.xx3.focus()
      document.add.xx3.select()
      return
     }
	 
  
 	 
  if (document.add.xx5.value=="")
     {
      alert("请填写联系电话！")
      document.add.xx5.focus()
      document.add.xx5.select()
      return
     }
	 	 
     document.add.submit()
}
-->
</script>

</head>

<body>
<div align="center">
<table width="778" border="0" cellpadding="0" bgcolor="#99CCFF">
  <tr>
    <td colspan="2"><!--#include file="top3.asp"--></td>
  </tr>
<!-- <tr>
    <td width="205" valign="top"> 
        			<li><a href="psy.asp" target="main">
					<span style="background-color: #C0C0C0">
					<font color="#000080">心理健康宣传</font></span></a> 
					<li><a href="psytest.asp" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">心理测试</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">心理咨询预约</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">查看心理咨询师信息</span></font></a></li>

                    <!--<li><a href="#" target="main">
					<font color="#000080">
					<span style="background-color: #C0C0C0">活动认证管理</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动</span></font></a></li>
					<li><a href="#" target="main"><font color="#000080">
					<span style="background-color: #C0C0C0">添加活动类</span></font></a></li>
					-->
					</td>
    <td width="778" valign="top">
<!--<table width="100%"   border="0" cellpadding="3" cellspacing="1" bgcolor="#B5CDE3">-->
 </table>
<table width="778" align="center" bgcolor="#F0F7FF">
<%if ks=1 then %>
 <tr>
  <td id="info" colspan="4">请填写完整的预约信息，然后点“提交预约”按扭。</td>
 </tr>
 <tr>
  <td id="info2" colspan="4"></td>
 </tr>
 <tr>
  <td colspan="4" style="line-height:35px">
   <table width="100%" align="center" bgcolor="#FAFAFA"> 
     <form name="add" method="post" action="submit.asp">
	  <tr bgcolor="#FFFFFF">
		  <td width="20%" align="right">预约日期：</td>
		  <td width="79%"><input name="xx1" type="text" id="xx1" value="<%=request("rq")%>" readonly> </td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>星　　期：</td>
		  <td width="79%"><input name="xx2" type="text" id="xx2" value="<%=request("xq")%>" readonly/></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>预约时间：</td>
		  <%
		  %>
		  <td width="79%"><input name="xx3" type="text" id="xx3"  value="第<%=request("ks")%>节" readonly/></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>预约医师：</td>
          <td width="79%">
		  <SELECT name='xx4' id="xx4">
		  <option value="<%=request("cid")%>"><%=request("cid")%></option>
		  </SELECT></td>

		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>联系电话：</td>
		  <td width="79%"><input name="xx5" type="text" id="xx5"></td>
		</tr>
		<tr bgcolor='#FFFFFF'>
		  <td align='right'>备注说明：</td>
		  <td width="79%"  ><textarea name="xx6" cols="53" rows="3" id="xx6"></textarea></td>
		</tr>
	    <tr bgcolor='#FFFFFF'>
		  <td align='right' >申请时间：</td>
		  <td width="79%" ><input name="addtime" type="text" id="addtime" value="<%=now()%>" readonly></td>
		 </tr>
        <tr align="center" bgcolor="#FFFFFF">
          <td colspan="2" ><input TYPE="hidden" name="action" value="yes"><input type="button" name="Submit" value="提交预约" onClick="check()"></td>
        </tr>
		</form>
      </table>
  </td>
 </tr>
<%end if%>
</table>


	</td>
  </tr>
  <tr>
    <td width="778" colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</div>
</body>
</html>