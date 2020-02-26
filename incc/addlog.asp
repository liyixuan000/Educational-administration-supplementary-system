<%function addlog(content)
sql="select * from Eptime_log"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
rs("ip")=request.servervariables("remote_addr")
rs("os")=browser(Request.ServerVariables("HTTP_USER_AGENT"))&"/"&system(Request.ServerVariables("HTTP_USER_AGENT"))
rs("content")=content
rs.update
rs.close
set rs=nothing
end Function

function addlog1(id)
sql="select * from Eptime_admin where id="&id&""
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
rs("LastLoginIP")=request.servervariables("remote_addr")
rs("LastLoginTime")=now()
rs.update
rs.close
set rs=nothing
end Function

function browser(info)
if Instr(info,"NetCaptor 6.5.0")>0 then
browser="NS 6.5.0"
ElseIf InStr(info,"GreenBrowser")>0 Then
browser="电脑客户端"
elseif Instr(info,"IQ")>0 then
browser="IQ"
elseIf InStr(info," Firefox/")>0 Then
browser="Mozilla Firefox"
ElseIf InStr(info," Firebird/")>0 Then
browser="Mozilla Firebird"
ElseIf InStr(info,"Maxthon")>0 Then
browser="Maxthon"
ElseIf InStr(info,"QQBrowser")>0 Then
browser="腾讯浏览器"
elseif Instr(info,"MyIe 3.1")>0 then
browser="遨游 3.1"
elseif Instr(info,"NetCaptor 6.5.0RC1")>0 then
browser="NS 6.5.0RC1"
elseif Instr(info,"NetCaptor 6.5.PB1")>0 then
browser="NS 6.5.PB1"
elseif Instr(info,"MSIE 5.5")>0 then
browser="IE 5.5"
elseif Instr(info,"MSIE 9.0")>0 then
browser="IE 9.0"
elseif Instr(info,"MSIE 8.0")>0 then
browser="IE 8.0"
elseif Instr(info,"MSIE 7.0")>0 then
browser="IE 7.0"
elseif Instr(info,"MSIE 6.0")>0 then
browser="IE 6.0"
elseif Instr(info,"MSIE 6.0b")>0 then
browser="IE 6.0b"
elseif Instr(info,"MSIE 5.01")>0 then
browser="IE 5.01"
elseif Instr(info,"MSIE 5.0")>0 then
browser="IE 5.00"
elseif Instr(info,"MSIE 4.0")>0 then
browser="IE 4.01"
else
browser="手机客户端"
end if
end function
function system(info)
    if Instr(info,"NT 5.1")>0 then
        system=system+"Windows XP"
    elseif Instr(info,"Tel")>0 then
        system=system+"Telport"
	elseif Instr(info,"webzip")>0 then
        system=system+"webzip"
	elseif Instr(info,"flashget")>0 then
        system=system+"flashget"
	elseif Instr(info,"offline")>0 then
        system=system+"offline"
    elseif Instr(info,"NT 6.2")>0 then
        system=system+"Windows 7"
    elseif Instr(info,"NT 6.1")>0 then
        system=system+"Windows 7"
    elseif Instr(info,"NT 6.0")>0 then
        system=system+"Windows vista"
    elseif Instr(info,"NT 5.2")>0 then
        system=system+"Windows 2003"
    elseif Instr(info,"NT 5")>0 then
        system=system+"Windows 2000"
    elseif Instr(info,"NT 4")>0 then
        system=system+"Windows NT4"
    elseif Instr(info,"98")>0 then
        system=system+"Windows 98"
    elseif Instr(info,"95")>0 then
        system=system+"Windows 95"
	elseif instr(info,"unix") or instr(info,"linux") or instr(info,"SunOS") or instr(info,"BSD") then
	    system=system+"类Unix"
    elseif instr(thesoft,"Mac") then
	    system=system+"Mac"
    elseif instr(thesoft,"Android") then
	    system=system+"APP"
    else
        system=system+"APP"
    end if
end function
%>