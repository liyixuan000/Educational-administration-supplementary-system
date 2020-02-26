<!-- #include file="incc/conn.asp" -->
<html>
<head>
<title>详细信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin.css" rel="stylesheet" type="text/css">
<style>
body {
	background-color:#FFFFFF;
}
</style>
<script language=javascript>
function preview() { 
bdhtml=window.document.body.innerHTML;
sprnstr="<!--startprint-->"; 
eprnstr="<!--endprint-->"; 
prnhtml=bdhtml.substr(bdhtml.indexOf(sprnstr)+17); 
prnhtml=prnhtml.substring(0,prnhtml.indexOf(eprnstr));
window.document.body.innerHTML=prnhtml; 
window.print();
window.document.body.innerHTML=bdhtml; 
         }
</script>

</HEAD>

<BODY>
<%
sql="select * from Eptime_money where id="&request("id")
set rs=conn.execute(sql)
%><br>

  <table cellpadding="0" cellspacing="0" width="96%">
    <tr>
	  <td align="right">&nbsp;<img src="images/print.jpg" align="absmiddle" style="cursor:hand;" onClick="preview();"></td>
    </tr>
  </table>

<table width="96%" border="0" cellpadding="0" cellspacing="0" bgcolor="#EBEBEB" align="center">
<tr>
<td><img src="images/r_1.gif" alt="" /></td>
<td width="100%" background="images/r_0.gif">
</td>
<td><img src="images/r_2.gif" alt="" /></td>
</tr>
<tr>
<td height="423"></td>
<td height="423">
<!--startprint-->
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr height="30">
      <td>&nbsp;单号：<font color="#ff0000"><%=rs("moneyID")%></font> 流水账详细信息</td>
	  <td align="right">　</td>
    </tr>
  </table>
<table align="center" cellpadding="4" cellspacing="1" class="toptable grid" border="1" width="100%">
	  <tr>
        <td align="right" height="30">时间：</td>
        <td class="category">
		  <%=rs("addtime")%></td>
      </tr>	
     	  <tr>
        <td align="right" height="30">所属大类：</td>
        <td class="category">
          <%
	      sql="select * from Eptime_money_bigclass where id="&rs("id_bigclass")
	      set rs_bigclass=conn.execute(sql)
	      %>
          <%if rs_bigclass.eof=false then%><%=rs_bigclass("bigclass")%><%else%>&nbsp;<%end if%>	</td>
      </tr>  
      <tr>
        <td align="right" height="30">所属小类：</td>
        <td class="category">
          <%
	      sql="select * from Eptime_money_smallclass where id="&rs("id_smallclass")
	      set rs_smallclass=conn.execute(sql)
		  %>
          <%if rs_smallclass.eof=false then%><%=rs_smallclass("smallclass")%><%else%>&nbsp;<%end if%></td>
      </tr>	   
      <tr>
        <td width="25%" height="30" align="right">金额： </td>
        <td width="75%" class="category"><%=rs("price")%></td>
      </tr>	
   	  
	
      <tr>
        <td align="right" height="30">业务项目：</td>
        <td class="category">
	  <%
	  sql="select * from Eptime_xiangmu where id="&rs("xiangmu")
	  set rs_xiangmu=conn.execute(sql)
	  %>
	  <%if rs_xiangmu.eof=false then%><%=rs_xiangmu("name")%><%else%>&nbsp;<%end if%>	
		</td>
      </tr>

           <tr>
        <td align="right" height="30">备注：</td>
        <td class="category">
		  <%=replace(replace(rs("beizhu")&""," ","&nbsp;"),chr(13),"<br>")%></td>
      </tr>	


      <tr>
        <td align="right" height="30">记账人：</td>
        <td class="category">
	  <%
	  sql="select * from Eptime_admin where id="&rs("id_login")
	  set rs_login=conn.execute(sql)
	  %>
	  <%if rs_login.eof then%><%=rs("login")%><%else%><%=rs_login("realname")%><%end if%>	
		</td>
      </tr>

	  	  <tr>
        <td align="right" height="30">添加时间：</td>
        <td class="category">
		  <%=rs("addtime")%></td>
      </tr>	
	  	  	  <tr>
        <td align="right" height="30">最后一次修改人：</td>
        <td class="category">
		  <%=rs("LastLogin")%></td>
      </tr>	
	  	  	  <tr>
        <td align="right" height="30">最后一次修改时间：</td>
        <td class="category">
		  <%=rs("LastTime")%></td>
      </tr>	
<%If iscut="1" then%>
	  <tr>
        <td align="right" height="30">审核状态：</td>
        <td class="category">
		  <%If rs("isok")=False then%><font color="#CC3300">未审</font><%else%>已审<%End If%></td>
      </tr>	
	  	 <tr>
        <td align="right" height="30">审核人：</td>
        <td class="category">
		  <%=rs("shenhe")%></td>
      </tr>	
	  	  <tr>
        <td align="right" height="30">审核时间：</td>
        <td class="category">
		  <%=rs("shenhetime")%></td>
      </tr>
<%End if%>
</table>
<!--endprint-->
</td>
<td height="423"></td>
</tr>
<tr>
<td><img src="images/r_4.gif" alt="" /></td>
<td></td>
<td><img src="images/r_3.gif" alt="" /></td>
</tr>
</table>

</body>
</html>