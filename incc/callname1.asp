<%
function calladmin(nameid)
sql_call1="select * from Eptime_admin where id="&nameid&""
set rs_call1=conn.execute(sql_call1) 
if rs_call1.eof Then
Else
calladmin=rs_call1("username")
end if	
rs_call1.close
set rs_call1=nothing
end Function

function callbigclass(nameid)
sql_call2="select * from Eptime_perClass where id="&nameid
set rs_call2=conn.execute(sql_call2) 
if rs_call2.eof Then
Else
callbigclass=rs_call2("bigclass")
end if	
rs_call2.close
set rs_call2=nothing
end Function

function callsmallclass(nameid)
sql_call3="select * from Eptime_perxingshi where id="&nameid
set rs_call3=conn.execute(sql_call3) 
if rs_call3.eof Then
Else
callsmallclass=rs_call3("smallclass")
end if	
rs_call3.close
set rs_call3=nothing
end Function

function callname(id,nameid)
If id=1 Then
sql_call="select * from eptime_wanglai where id="&nameid&""
ElseIf id=2 Then
sql_call="select * from eptime_yuangong where id="&nameid&""
ElseIf id=3 Then
sql_call="select * from eptime_xiangmu where id="&nameid&""
ElseIf id=0 Then
sql_call="select * from eptime_zhanghu where id="&nameid&""
End If
set rs_call=conn.execute(sql_call) 
if rs_call.eof Then
Else
callname=rs_call("name")
end if	
rs_call.close
set rs_call=nothing
end Function

function selectname(id1,name1)
If id1=1 Then
sql_select="select * from eptime_wanglai order by id"
ElseIf id1=2 Then
sql_select="select * from eptime_yuangong order by id"
ElseIf id1=3 Then
sql_select="select * from eptime_xiangmu order by id"
ElseIf id1=0 Then
sql_select="select * from eptime_zhanghu order by id"
End If

set rs_select=conn.execute(sql_select)	  
if rs_select.eof then
Else
response.Write "<select name="&name1&">"	
          do while rs_select.eof=False
response.Write "<option value='"&rs_select("id")&"'>"&rs_select("name")&"</option>"
		    rs_select.movenext
		  Loop
response.Write "</select>"		  
end if	
rs_select.close
set rs_select=nothing
end Function


function selectname1(id1,name,id2)
If id1=1 Then
sql_select="select * from eptime_wanglai order by id"
ElseIf id1=2 Then
sql_select="select * from eptime_yuangong order by id"
ElseIf id1=3 Then
sql_select="select * from eptime_xiangmu order by id"
ElseIf id1=0 Then
sql_select="select * from eptime_zhanghu order by id"
End If

set rs_select=conn.execute(sql_select)	  
if rs_select.eof then
Else
response.Write "<select name='"&name&"'>"	
          do while rs_select.eof=False
If rs_select("id")=id2 Then
response.Write "<option value='"&rs_select("id")&"' selected='selected'>"&rs_select("name")&"</option>"
Else
response.Write "<option value='"&rs_select("id")&"' >"&rs_select("name")&"</option>"
End If 
		    rs_select.movenext
		  Loop
response.Write "</select>"		  
end If
rs_select.close
set rs_select=nothing
end Function
%>