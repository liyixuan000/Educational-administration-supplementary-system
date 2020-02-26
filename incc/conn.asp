<!-- #include file="config.asp" -->
<%
If(databaseType=0) Then 
'ACCESS数据库
set conn=server.createobject("adodb.connection")
mypath=server.mappath("Database/db_database.mdb")
conn.open "provider=microsoft.jet.oledb.4.0;data source=" & mypath
ElseIf(databaseType=1) Then 
'MSSQL SERVER数据库
    Set Conn=Server.CreateObject("Adodb.Connection") 
    StrConn = "PROVIDER=SQLOLEDB.1;Data Source="&databaseServer&";Initial Catalog="&databaseName&";Persist Security Info=True;User ID="&databaseUser&";Password="&databasePwd&";Connect Timeout=30" 
    Conn.Open StrConn 
Else 
'数据库设置错误
    Response.Write"数据库设置错误，请联系管理员！" 
    Response.End 
End If 
'Response.Write StrConn

'只取数字
Function Trim2(ke2)
   Set ob2=New Regexp
   ob2.IgnoreCase=True
   ob2.Global=True
   ob2.Pattern="[a-zA-Z]+"
   Trim2=ob2.Replace(ke2,"")
   Set ob2=Nothing
End Function

'截取内容，从“”开始，到“”结束
Function GetKey(HTML,Start,Last) 
filearray=split(HTML,Start) 
filearray2=split(filearray(1),Last) 
GetKey=filearray2(0) 
End Function 

'url字符反编译
function URLDecode(enStr)
dim deStr,strSpecial
dim c,i,v
  deStr=""
  strSpecial="!""#$%&'()*+,.-_/:;<=>?@[\]^`{|}~%"
  for i=1 to len(enStr)
    c=Mid(enStr,i,1)
    if c="%" then
      v=eval("&h"+Mid(enStr,i+1,2))
      if inStr(strSpecial,chr(v))>0 then
        deStr=deStr&chr(v)
        i=i+2
      else
        v=eval("&h"+ Mid(enStr,i+1,2) + Mid(enStr,i+4,2))
        deStr=deStr & chr(v)
        i=i+5
      end if
    else
      if c="+" then
        deStr=deStr&" "
      else
        deStr=deStr&c
      end if
    end if
  next
  URLDecode=deStr
end Function

function makefilename(danhao)
makefilename=""
t=now()
makefilename=year(t)
if month(t)<10 then makefilename=makefilename&"0"
makefilename=makefilename&month(t)
if day(t)<10 then makefilename=makefilename&"0"
makefilename=makefilename&day(t)
if hour(t)<10 then makefilename=makefilename&"0"
makefilename=makefilename&hour(t)
if minute(t)<10 then makefilename=makefilename&"0"
makefilename=makefilename&minute(t)
if second(t)<10 then makefilename=makefilename&"0"
makefilename=makefilename&second(t)
If request.Cookies("eptime_id")<10 Then makefilename=makefilename&"0"
makefilename=makefilename&request.Cookies("eptime_id")
makefilename=danhao&makefilename
end Function

Function Format_Time(s_Time, n_Flag)
 Dim y, m, d, h, mi, s
 Format_Time = ""
 If IsDate(s_Time) = False Then Exit Function
 y = cstr(year(s_Time))
 m = cstr(month(s_Time))
 If len(m) = 1 Then m = "0" & m
 d = cstr(day(s_Time))
 If len(d) = 1 Then d = "0" & d
 h = cstr(hour(s_Time))
 If len(h) = 1 Then h = "0" & h
 mi = cstr(minute(s_Time))
 If len(mi) = 1 Then mi = "0" & mi
 s = cstr(second(s_Time))
 If len(s) = 1 Then s = "0" & s
 Select Case n_Flag
 Case 1
 ' yyyy-mm-dd hh:mm:ss
 Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
 Case 2
 ' yyyy-mm-dd
 Format_Time = y & "-" & m & "-" & d
 Case 3
 ' hh:mm:ss
 Format_Time = h & ":" & mi & ":" & s
 Case 4
 ' yyyy年mm月dd日
 Format_Time = y & "年" & m & "月" & d & "日"
 Case 5
 ' yyyymmdd
 Format_Time = y & m & d
 case 6
 'yyyymmddhhmmss
 format_time= y & m & d & h & mi & s
 End Select
End Function
%><!--#include file="SF_Sql.asp"-->