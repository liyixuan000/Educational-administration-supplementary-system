<%
set rs2=server.createobject("adodb.recordset") 
sql2="select * from Site where id=1"
rs2.open sql2,conn,1,1
if not rs2.eof Then
SiteName=rs2("SiteName")
ks=rs2("Ys")
Description=rs2("Description")
Copyright=rs2("Copyright")
end if
rs2.close
set rs2=nothing
%>