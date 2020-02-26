<%
db="data/Xiao5u.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};pwd=;dbq=" & Server.MapPath(db) 
%>