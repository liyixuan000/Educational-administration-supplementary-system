<%
set conn=server.CreateObject("adodb.connection")
sql="Driver={Microsoft Access Driver (*.mdb)};DBQ= "&server.mappath("Database/db_database.mdb")
conn.open(sql)
%>

