<%
'---------- ��ֹSQLע�� -----------
dim SQL_Injdata 
SQL_Injdata = "'|;|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare" 
SQL_inj = split(SQL_Injdata,"|") 
If Request.QueryString<>"" Then 
  For Each SQL_Get In Request.QueryString 
    For SQL_Data=0 To Ubound(SQL_inj) 
      if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_Data))>0 Then 
        Response.Redirect("/index.asp")
      end if 
    next 
  Next 
End If
'---------- �������ݿ� ----------
Dim conn,connstr
Set conn=Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Password=;Data Source="&Server.MapPath("DataBase/db_database.mdb")&";"
conn.open connstr
'---------- ͳ�Ʒ����� ----------
If Trim(Request.Cookies("visitor"))="" Then
'��ȡ�ļ�count.txt������
Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
Set TextFile=FSObject.OpenTextFile(Server.MapPath("count.txt"))
If not TextFile.AtEndOfStream Then
num3=TextFile.ReadLine
End If
Set TextFile=Nothing
'���ļ�count.txt��׷������
Set TextFile=FSObject.OpenTextFile(Server.MapPath("count.txt"),2,true)
TextFile.WriteLine num3+1
Set TextFile=Nothing
Set FSObject=Nothing
Response.Cookies("visitor")="visited"
Response.Cookies("visitor").Expires=DateAdd("d",1,now()) 'ָ��Cookie����Чʱ��
End If

Function saferequest(value)
	Dim ParaValue
	ParaValue = Trim(Request(value))
	If IsNumeric(ParaValue) = True Then
		saferequest = ParaValue
		Exit Function
	
	ElseIf InStr(LCase(ParaValue), "select ") > 0 Or InStr(LCase(ParaValue), "insert ") > 0 Or InStr(LCase(ParaValue), "delete from") > 0 Or InStr(LCase(ParaValue), "count(") > 0 Or InStr(LCase(ParaValue), "drop table") > 0 Or InStr(LCase(ParaValue), "update ") > 0 Or InStr(LCase(ParaValue), "truncate ") > 0 Or InStr(LCase(ParaValue), "asc(") > 0 Or InStr(LCase(ParaValue), "mid(") > 0 Or InStr(LCase(ParaValue), "char(") > 0 Or InStr(LCase(ParaValue), "xp_cmdshell") > 0 Or InStr(LCase(ParaValue), "exec master") > 0 Or InStr(LCase(ParaValue), "net localgroup administrators") > 0 Or InStr(LCase(ParaValue), " and ") > 0 Or InStr(LCase(ParaValue), "net user") > 0 Or InStr(LCase(ParaValue), " or ") > 0 Or InStr(LCase(ParaValue), "'") > 0 Or InStr(LCase(ParaValue), "''") > 0 Then
		response.Write("<script>alert('�벻Ҫ����Ƿ��ַ�');history.back();</script>")
		response.End()  
	Else
		saferequest = ParaValue
	End If
End Function
dim lan_txt:lan_txt="gb"
%>
