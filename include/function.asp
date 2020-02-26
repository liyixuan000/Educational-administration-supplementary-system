<%
'---------- 定义转换字符函数 -----------
Function Str_filter(InString)
  NewStr=Replace(InString,"'","''")
  NewStr=Replace(NewStr,"<","&lt")
  NewStr=Replace(NewStr,">","&gt")
  NewStr=Replace(NewStr,"chr(60)","&lt;")
  NewStr=Replace(NewStr,"chr(37)","&gt;")
  NewStr=Replace(NewStr,"""","&quot")
  NewStr=Replace(NewStr,";",";;")
  NewStr=Replace(NewStr,"--","-")
  NewStr=Replace(NewStr,"/*"," ")
  NewStr=Replace(NewStr,"%"," ")
  Str_filter=NewStr
End Function
'---------- 根据时间获取的字符串 -----------
Function GetOrderNo(dDate)
    GetOrderNo = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)+RIGHT("00" + Trim(Hour(dDate)),2)+RIGHT("00"+Trim(Minute(dDate)),2)+RIGHT("00"+Trim(Second(dDate)),2)
End Function
'---------- 获取随机数函数 -----------
Function randStr(num)
  strings="0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
  str=split(strings,",")
  for i=1 to num
  Randomize
  str1=str(int((ubound(str))*rnd))
  returnstr=returnstr&str1
  next
  randStr=returnstr
End Function

Function Get_Tree(Fatherid,Fatherid_Old)
	Dim Sqlstr,Rs,i
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	Sqlstr = "Select A_id,A_MenuName,A_Level From B_adminmenu  Where A_parentid="&Fatherid&" Order by A_Orderby"
	Rs.Open Sqlstr,Conn,1,1
	Do While Not Rs.Eof
		If Fatherid_Old=Rs("A_id") Then
			Response.Write("<option value='"&Rs("A_id")&"' selected>")
		Else
			Response.Write("<option value='"&Rs("A_id")&"'>")
		End If
		For i=0 to RS("A_Level")-1
			Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
		Next
		Response.Write(Rs("A_MenuName")&"</option>")
		Call Get_Tree(Rs("A_id"),Fatherid_Old)
		Rs.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
End Function
'a为要去字符的内容,b为开始字符,c为结束字符   去除从某段字符到另一字符之间的字符串  
Function StrChar(a,b,c)
Dim ConStrTemp,StartStr,OverStr,Start,Over,GetBody,d
   do while instr(a,b)>0 and instr(a,c)>0
	   ConStrTemp=Lcase(a)
	   StartStr=Lcase(b)
	   OverStr=Lcase(c)
	   Start = InStrB(1, ConStrTemp, StartStr, vbBinaryCompare)
	   Start = Start+LenB(StartStr)
	   Over = InStrB(Start,ConStrTemp,OverStr,vbBinaryCompare)
	   Over = Over+LenB(OverStr)
	   GetBody = MidB(a,Start,Over-Start)
	   d=b&GetBody
	   a=replace(a,d,"")
   loop
   StrChar=a
end function

Function HTMLEncode(ByVal reString) 
	Dim Str:Str=reString
	If Not IsNull(Str) Then
		Str = Replace(Str, CHR(10), "<br/>")
   		Str = Replace(Str, ">", "&gt;")
		Str = Replace(Str, "<", "&lt;")
	    Str = Replace(Str, CHR(9), "&nbsp;")
	    Str = Replace(Str, CHR(39), "&#39;")
    	Str = Replace(Str, CHR(34), "&quot;")
		Str = Replace(Str, CHR(13), "")
		HTMLEncode = Str
	End If
End Function

Function HTMLEncode_(ByVal reString) 
	Dim Str:Str=reString
	If Not IsNull(Str) Then
   		Str = Replace(Str, "&gt;", ">")
		Str = Replace(Str, "&lt;", "<")
	    Str = Replace(Str, "&nbsp;", CHR(9))
	    Str = Replace(Str, "&#39;", CHR(39))
    	Str = Replace(Str, "&quot;", CHR(34))
		Str = Replace(Str, "", CHR(13))
		Str = Replace(Str, "<br/>", CHR(10))
		HTMLEncode_ = Str
	End If
End Function

function checkint(str,def)
   '检测输入的是否是整数
   'str 输入的字符串，def如果str非法则返回的整数
   if len(str)= 0 or isnull(str) then
       checkint = def
       exit function
   end if
   if isnumeric(str) then
       checkint=clng(str)
   else
       checkint=def
   end if
end function

'================================================
'函数名：GetStrLen
'作  用：'取字符串长度，中文为两个字节
'参  数：str   ----原字符串
'================================================
Function GetStrLen(Str) 
	Dim P_Len, XX 
	P_Len = 0 
	GetStrLen = 0 
	If Not IsNull(Str) And Str <> "" Then 
		P_Len = Len(Str) 
		For XX = 1 To P_Len 
			If Asc(Mid(Str, XX, 1)) < 0 Then 
				GetStrLen = CLng(GetStrLen) + 2 
			Else 
				GetStrLen = CLng(GetStrLen) + 1 
			End If 
		Next 
	End If 
End Function


'=============================================================
'函数作用：检测身份证号码是否正确
'=============================================================
Function CheckCardId(e)   
    arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")   
    Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")   
    Checker = Split("1,9,8,7,6,5,4,3,2,1,1", ",")   
    If Len(e) < 15 Or Len(e) = 16 Or Len(e) = 17 Or Len(e) > 18 Then   
        CheckCardId= "身份证号共有 15 码或18位"     
        Exit Function   
    End If   
    Dim Ai   
    If Len(e) = 18 Then   
        Ai = Mid(e, 1, 17)   
    ElseIf Len(e) = 15 Then   
        Ai = e   
        Ai = Left(Ai, 6) & "19" & Mid(Ai, 7, 9)   
    End If   
    If Not IsNumeric(Ai) Then   
        CheckCardId= "身份证除最后一位外，必须为数字！"           
        Exit Function   
    End If   
    Dim strYear, strMonth, strDay   
    strYear = CInt(Mid(Ai, 7, 4))   
    strMonth = CInt(Mid(Ai, 11, 2))   
    strDay = CInt(Mid(Ai, 13, 2))   
    BirthDay = Trim(strYear) + "-" + Trim(strMonth) + "-" + Trim(strDay)   
    If IsDate(BirthDay) Then   
        If DateDiff("yyyy",Now,BirthDay)<-140 or cdate(BirthDay)>date() Then           
            CheckCardId= "身份证输入错误！"   
            Exit Function   
        End If   
        If strMonth > 12 Or strDay > 31 Then   
            CheckCardId= "身份证输入错误！"   
            Exit Function   
        End If   
    Else   
        CheckCardId= "身份证输入错误！"   
        Exit Function   
    End If   
    Dim i, TotalmulAiWi   
    For i = 0 To 16   
        TotalmulAiWi = TotalmulAiWi + CInt(Mid(Ai, i + 1, 1)) * Wi(i)   
    Next   
    Dim modValue   
    modValue = TotalmulAiWi Mod 11   
    Dim strVerifyCode   
    strVerifyCode = arrVerifyCode(modValue)   
    Ai = Ai & strVerifyCode    
    CheckCardId = Ai  
    If Len(e) = 18 And e <> Ai Then 
		CheckCardId= "身份证输入错误！"   
        Exit Function   
    End If 
	CheckCardId ="True"  
End Function 

'=============================================================
'函数作用：匹配正则表达式
'=============================================================
Function CheckExp(patrn, strng) 
	Dim regEx,Match'建立变量。 
	Set regEx = New RegExp'建立正则表达式。 
	regEx.Pattern = patrn'设置模式。 
	regEx.IgnoreCase = true'设置是否区分字符大小写。 
	regEx.Global = True'设置全局可用性。 
	Matches = regEx.test(strng)'执行搜索。 
	CheckExp = matches
End Function

'-------------------------------------------------------------
' Function Name : IsValidString
' Function Desc : 判断输入是否是一个由 0-9 / A-Z / a-z / . / _ 组成的字符串
'-------------------------------------------------------------
Function IsValidString(sInput)
	
	Dim oRegExp
	
	'建立正则表达式
	Set oRegExp = New RegExp
	
	'设置模式
	oRegExp.Pattern = "^[a-zA-Z0-9\s.\-_]+$"
	'设置是否区分字符大小写
	oRegExp.IgnoreCase = True
	'设置全局可用性
	oRegExp.Global = True
	
	'执行搜索
	IsValidString = oRegExp.Test(sInput)
	
	Set oRegExp = Nothing

End Function
	
function checknumber(str,def)
	if len(str)= 0 or isnull(str) then
       checknumber = def
       exit function
   end if
   if isnumeric(str) then
       checknumber=str
   else
       checknumber=def
   end if
end function

'安全除法
Function SafeDiv(A,B)
	SafeDiv = checknumber(A,0) / checknumber(B,1)
End Function

function checksqlstr(getstr)
'检测输入的参数是否含有sql敏感字符，如果有返回空字符串
	dim strfilter,strtmp,i,regEx
	if len(getstr) = 0 or isnull(getstr) then
		checksqlstr = ""
		exit function
	end if
	Set regEx = New RegExp
	strfilter = "select|delete|update|drop|create|exec"
	regEx.Pattern =	strfilter
	regEx.IgnoreCase = True
	regEx.Global = True
	getstr = trim(regex.Replace(getstr,""))
	strfilter="'"
	regEx.Pattern=   strfilter
	getstr = trim(regex.Replace(getstr,"''"))
	
	strfilter="0x"
	regEx.Pattern=   strfilter
	getstr = trim(regex.Replace(getstr,""))
	
	
	regex.Pattern="法[\s　]*轮[\s　]*功"
	getstr=regex.Replace(getstr,"*轮*")
	set regex=nothing
	checksqlstr = getstr
end function

function checkuserstr(str)
   '检测用户注册名
	if len(str) = 0 or isnull(str) then
		checkuserstr = ""
		exit function
	end if
	dim i 
	dim lb_user,strfilter,regEx,str1
	str1 = "[A-Za-z0-9|\u4e00-\u9fa5]*" '输入的数据必须是中文
	strfilter="\$|\(|\)|\*|\+|\-|\.|\[|]|\?|\\|\^|\{|\||}|~|`|!|@|#|%|&|_|=|<|>|/|,|'| |　|\:|cubcn|admin|administrator|administrators|luo|老罗|system"
	Set regEx = New RegExp
	regEx.Pattern =	strfilter
	regEx.IgnoreCase = True
	lb_user = regex.Test(str)
	if not lb_user then
		regex.Pattern="法[ 　]*轮[ 　]*功"
		checkuserstr = regex.Replace(str,"")
	else
		checkuserstr = ""
	end if
	
	'regEx.Pattern =	str1
	'regEx.IgnoreCase = True
	'lb_user = regex.Test(str)
	'if not lb_user then		
	'	checkuserstr = ""
	'end if
	
	'dim regEx2,strfilter1
	'strfilter1 = "^[A-Za-z0-9\u4e00-\u9fa5]+$" '输入的数据必须是中文
	'Set regEx2 = New RegExp
	'regEx2.Pattern =	strfilter1
	'regEx2.IgnoreCase = True
	
	'if not regex2.Test(str) then		
	'	checkuserstr = ""
	'end if
	'set regex2 = nothing
	
	set regex=nothing	
end function 



'组合文件路径
function JoinPath(Fstr,Estr)
	Dim T_Fstr,T_Estr,TmpStr
	T_Fstr = Fstr
	T_Estr = Estr
	TmpStr=Right(Fstr,1)
	if  TmpStr="/" or TmpStr ="\" then
		T_Fstr = Left(Fstr,len(Fstr)-1)
	end if
	TmpStr=Left(Estr,1)
	if TmpStr="/" or TmpStr ="\" then
		T_Estr = Right(Estr,len(Estr)-1)
	end if
	JoinPath = T_Fstr & "/" & T_Estr
end function

'把虚拟路径,转志当前物理路径
Function PathToCurrent(FpathStr)
	PathToCurrent = Server.MapPath(FpathStr)
End Function


'返回文件名的后缀 参数是文件名
Function ReturnFileExt(SouStr)
	Dim Ind
	Ind = InStr(SouStr,".")
	ReturnFileExt = Mid(SouStr,Ind)
End Function

Function Cutstr(str,num)
	If GetStrLen(str)>num Then 
		str = Left(str,num/2)
	End If
	Cutstr = str
End Function

Function CutstrPoint(str,num)
	If GetStrLen(str)>num Then 
		str = Left(str,num/2) & "..."
	End If
	CutstrPoint = str
End Function

Sub jstop(strmsg)
'显示信息并回退一步
	dim html
	html = "<script>alert('" & strmsg & "');if(history.length==0){window.opener='';window.close();}else{history.back();}</script>"
	response.Write html
	'response.redirect "../index.asp"
	response.End 
End sub

Sub JsAlertMsg(strmsg)
	Dim html
	html = "<script>alert('" & strmsg & "');</script>"
	response.Write html
End Sub

Sub JsAlertAndGoto(strmsg,Url)
	Dim html
	html = "<script>alert('" & strmsg & "');window.location.href='"& Url&"';</script>"
	response.Write html
	Response.end
End Sub

Sub JsAlertAndGoto_1(strmsg,Url01,Url02)
	Dim html
	html = "<script>if(confirm('" & strmsg & "')){window.location.href='"& Url01&"';}else{window.location.href='"& Url02&"';}</script>"
	response.Write html
	Response.end
End Sub

Sub JsAlertAndGoto_(strmsg,Url)
	Dim html
	html = "<script>alert('" & strmsg & "');window.open('"& Url&"','_blank');</script>"
	response.Write html
	Response.end
End Sub


Rem 得到内容
Function GetTxt(Path)
	On Error resume next
	Dim Ph,fso,ts
	Set fso = CreateObject("Scripting.FileSystemObject")
	ph=server.mappath(Path)
	Set ts = fso.OpenTextFile(ph, 1)
	GetTxt = ts.ReadAll
	ts.Close
End Function

Rem 检测目录是否存在,不存在则建立目录
Function CreateFolder(Folder)
	DIM FSO
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if Not fso.FolderExists(Server.MapPath(Folder)) Then 
		fso.CreateFolder(Server.MapPath(Folder))
	End if
	Set Fso = Nothing
	If Err.number<>0 Then
	 err.Clear
	  Response.Write "权限错误!!请检查您的IIS默认的匿名用户是否对本系统所在的目录有读,写,修改权限!!2"
	  Response.End
	End If
End Function
Rem 删除文件
Function DelFile(Path)
	DIM fso
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	IF fso.FileExists(Server.MapPath(Path)) then
		fso.DeleteFile(Server.MapPath(Path))
	End IF
	Set fso=Nothing
End Function


Rem 删除目录
Function DelFolder(Path)
	DIM fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	IF fso.FolderExists(Server.MapPath(Path)) then
		fso.DeleteFolder(Server.MapPath(Path))
	End IF
	Set fso=Nothing
End Function


Rem 拷贝文件
Sub copyfile(movefiles, copyfiles)
	movefiles = Server.MapPath(movefiles)
	copyfiles = Server.MapPath(copyfiles)
	Dim Fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.CopyFile movefiles, copyfiles

	Set fso = Nothing
End Sub

Rem 检测用户上传目录是否已满
Function IsFolder(UserID,Zise)
	IsFolder=0
	Dim IsFolderf,IsFolderfso
	Set IsFolderfso = CreateObject("Scripting.FileSystemObject") 
	Set IsFolderf = IsFolderfso.GetFolder(Server.MapPath(GetUpLoadDir(UserID)))
	IsFolder=IsFolderf.Size 
	Set IsFolderf=Nothing:Set IsFolderfso=Nothing
	IF IsFolder>=(Zise*1024*1024) Then
		IsFolder=False
	Else
		IsFolder=True
	End IF
End Function


REM 写入文件
Sub WriteStream(FileName,BodyText) 
	On Error resume next
	DIM WriteAdo
	Set WriteAdo = Server.CreateObject("ADODB.Stream")
	With WriteAdo
		.Open
		.Charset = "gb2312"
		.WriteText BodyText
		.SaveToFile Server.MapPath(FileName),2
		.Close
	End With
	Set WriteAdo = Nothing
End Sub


Function makefilename(fname)
  'fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"上午","")
  fname = replace(fname,"下午","")
  makefilename=fname
end Function

'===================================================================
'                           时间转换函数2
'===================================================================
function chan_data(shijian)
	Dim s_year,s_month,s_day,s_hour
	s_year=year(shijian)
	if len(s_year)=2 then s_year="20"&s_year
	s_month=month(shijian)
	if s_month<10 then s_month="0"&s_month
	s_day=day(shijian)
	if s_day<10 then s_day="0"&s_day
	s_hour=hour(shijian)
	chan_data =s_year & s_month 
	
end function
'===================================================================
'                           时间转换函数3
'===================================================================
function chan_day(shijian)	
	Dim s_day
	s_day=day(shijian)	
	chan_day=s_day 
	
end function

'===================================================================
'函数功能：替换文章内容模版内容
'参数说明：(略)
'===================================================================
Function Replace_ModeContent(TempletContent,SiteTitle,SiteUrl,position,title,keywords,Addtime,Sources,author,Writer,content,id,classid,ClassName,Join_Record) 
	TempletContent=replace(TempletContent,"$SiteTitle$",SiteTitle)
	TempletContent=replace(TempletContent,"$SiteUrl$",SiteUrl)
	TempletContent=replace(TempletContent,"$Position$",position)
	TempletContent=replace(TempletContent,"$Title$",title)
	TempletContent=replace(TempletContent,"$keywords$",keywords)
	TempletContent=replace(TempletContent,"$Addtime$",Addtime)
	TempletContent=replace(TempletContent,"$Source$",Sources)
	TempletContent=replace(TempletContent,"$Author$",author)
	TempletContent=replace(TempletContent,"$Writer$",Writer)
	TempletContent=replace(TempletContent,"$Content$",content)
	TempletContent=replace(TempletContent,"$Id$",id)
	TempletContent=replace(TempletContent,"$Classid$",classid)	
	TempletContent=replace(TempletContent,"$ClassName$",ClassName)
	TempletContent=replace(TempletContent,"$Join_Record$",Join_Record)		'相关记录
End Function	

'***********************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'***********************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'***********************************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<form name='showpages' method='Post' action='" & sfilename & "' style='margin:0 0 0 0;'><table align='center'  border='0' cellspacing='0' cellpadding='0'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table></form>"
	response.write strTemp
End Sub

'***********************************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************
sub showpageback(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='center'  border='0' cellspacing='0' cellpadding='0'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
End Sub

'***********************************************
'过程名：Show_Page
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************
Sub Show_Page(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit,CurrentPage)
	dim AllPage,i,ShowText
	if totalnumber mod maxperpage=0 then
    	AllPage = totalnumber \ maxperpage
  	else
    	AllPage = totalnumber \ maxperpage+1
  	end if
	if sfilename="" then
		TheUrl="?page="
	else
		TheUrl=sfilename&"&page="
	end if
	ShowCount=4
	if AllPage>8 then
		If (AllPage-CurrentPage)>=ShowCount and CurrentPage>0 then
			if CurrentPage<=ShowCount then
				for i=1 to ShowCount+1
					If i<>CurrentPage Then
						ShowText=ShowText & "<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
					Else
						ShowText=ShowText & "<LI class=pagebarCurrent>"&i&"</LI>"
					End If
				Next
			else
				If (CurrentPage-1)>0 then
					for i=(CurrentPage-1) to (CurrentPage+ShowCount-1-2)
						If i<>CurrentPage Then
							ShowText=ShowText & "<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
						Else
							ShowText=ShowText & "<LI class=pagebarCurrent>"&i&"</LI>"
						End If
					Next
				Else
					for i=CurrentPage to (CurrentPage+ShowCount-2)
						If i<>CurrentPage Then
							ShowText=ShowText & "<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
						Else
							ShowText=ShowText & "<LI class=pagebarCurrent>"&i&"</LI>"
						End If
					Next
				End If
			end if
		Else
			If (AllPage-ShowCount)>0 Then
				for i=AllPage-ShowCount to AllPage
					If i<>CurrentPage Then
						ShowText=ShowText & "<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
					Else
						ShowText=ShowText & "<LI class=pagebarCurrent>"&i&"</LI>"
					End If
				Next
			Else
				for i=1 to AllPage
					If i<>CurrentPage Then
						ShowText=ShowText & "<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
					Else
						ShowText=ShowText & "<LI class=pagebarCurrent>"&i&"</LI>"
					End If
				Next
			End If
		End If
		dim FirstText
		if CurrentPage>ShowCount and AllPage>5 then
			for i=1 to 2
				FirstText=FirstText&"<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
			next
			FirstText=FirstText&"<LI class=pagebarDot>... </LI>"
		end if
		if (AllPage-CurrentPage)>=ShowCount and  AllPage>5 then
			for i=AllPage-1 to AllPage
				EndText=EndText&"<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
			next
			EndText="<LI class=pagebarDot>... </LI>"&EndText
		end if
	elseif AllPage<=8 and AllPage>0 then
		for i=1 to AllPage
			If i<>CurrentPage Then
				ShowText=ShowText & "<LI><A  href='"&TheUrl&i&"' >"&i&"</A></LI>"
			Else
				ShowText=ShowText & "<LI class=pagebarCurrent>"&i&"</LI>"
			End If
		next
	end if
	
	ShowText=FirstText&ShowText&EndText
	if CurrentPage-1>0 then
		ShowText="<LI class=pagebarPrv><A href='"&TheUrl&CurrentPage-1&"'></A></LI>"&ShowText
	end if
	if CurrentPage+1<=AllPage then
		ShowText=ShowText&"<LI class=pagebarNext><A href='"&TheUrl&CurrentPage+1&"'></A></LI>"
	End if

	ShowText="<DIV class='nav pagebar fix'><UL>"&ShowText & "</ul></div>"
	if AllPage>1 then
		response.Write ShowText
	end if
End Sub

'**************************************************
'函数名：infotime
'作  用：格式年月日
'**************************************************
function infotime(a)
	dim monthstr,daystr,datestr
	monthstr="" :daystr=""
	if len(month(a))=1 then 
		monthstr="0"&month(a)
	else
		monthstr=month(a)
	end if
	if len(day(a))=1 then 
		daystr="0"&day(a)
	else
		daystr=day(a)
	end if
	datestr=year(a)&"-"&monthstr&"-"&daystr
	response.write(datestr)
end function

'**************************************************
'函数名：infotime_F,带字符
'作  用：格式年月日
'**************************************************
function infotimeStr(datetime,flag)
	dim monthstr,daystr,datestr
	monthstr="" :daystr=""
	if len(month(datetime))=1 then 
		monthstr="0"&month(datetime)
	else
		monthstr=month(datetime)
	end if
	if len(day(datetime))=1 then 
		daystr="0"&day(datetime)
	else
		daystr=day(datetime)
	end if
	datestr=year(datetime)&flag&monthstr&flag&daystr
	response.write(datestr)
end function

Sub CheckOutUrl()
	On Error Resume Next
	Dim server_v1, server_v2
	server_v1 = Replace(LCase(Trim(Request.ServerVariables("HTTP_REFERER"))), "http://", "")
	server_v2 = LCase(Trim(Request.ServerVariables("SERVER_NAME")))
	If server_v1 <> "" And Left(server_v1, Len(server_v2)) <> server_v2 Then
		Response.Write("请不要从外网跨站提交数据到本站！！！")
		Response.End()
	End If
End Sub

Function BuildFile(ByVal sFile, ByVal sContent)
On Error Resume Next
	Dim oFSO, oStream
	If CacheConfig(24) = 1 Then
		Set oFSO = server.CreateObject(CacheCompont(1))
'			If Is_Debug=1 Then Response.Write sFile
'			If Is_Debug=1 Then Response.Write sContent
		Set oStream = oFSO.CreateTextFile(sFile,True)
		oStream.Write sContent
		oStream.Close
		'增加对特殊字符的保护
		If Err.Number<>0 Then
			Set oStream = server.CreateObject(CacheCompont(2))
			With oStream
				.Type = 2
				.Mode = 3
				.open
				'.Charset = "gb2312"
				.Charset = "gb2312"
				.Position = oStream.size
				.WriteText = sContent
				.SaveToFile sFile, 2
				.Close
			End With
			Err.Clear
		End If
		Set oStream = Nothing
		Set oFSO = Nothing
	Else
		Set oStream = server.CreateObject(CacheCompont(2))
		With oStream
			.Type = 2
			.Mode = 3
			.open
			'.Charset = "gb2312"
			.Charset = "gb2312"
			.Position = oStream.size
			.WriteText = sContent
			.SaveToFile sFile, 2
			.Close
		End With
		Set oStream = Nothing
	End If
End Function


'取网页数据
Function GetHttp(url) 
 Dim Retrieval,GetBody
  Set Retrieval = Server.CreateObject("MSXML2.serverXMLHTTP") '//把单词拆开防止杀毒软件误杀
  Retrieval.Open "GET",url,false
  Retrieval.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
  Retrieval.Send
  GetBody=Retrieval.ResponseBody 
  GetHttp=BytesToBstr(GetBody)
  Set Retrieval = Nothing 
End function


'二进制转文本
Function BytesToBstr(body)
  Dim objstream
   Set objstream = CreateObject("Ado" & "db.Str" & "eam") '//把单词拆开防止杀毒软件误杀
   objstream.Type = 1
   objstream.Mode =3
   objstream.Open
   objstream.Write body
   objstream.Position = 0
   objstream.Type = 2
   objstream.Charset = "GB2312"
   BytesToBstr = objstream.ReadText 
   objstream.Close
 set objstream = Nothing
End function 
'------------------------WriteToHtml--------------------
Function WriteToHtml(Fpath,Templet)
	dim fs,f
	'On Error Resume Next 
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set f = fs.CreateTextFile(Fpath, True)

	f.write(Templet)

	f.close
	Set f = nothing
	Set fs = nothing
	if err <> 0 then
		WriteToHtml = false 
	else 
	  'call ConnectionDatabase()
		'sql="update tcarindex set turl='"&dbUrl&"' where CarID="&CarID
		'conn.execute(sql)
		'sql="update tcarweb set turl='"&dbUrl&"' where carid="&CarID
		'conn.execute(sql)
		'call closeConn()
		WriteToHtml = true 
	end if
End function

'用CDOnts组件发邮件
Function SendCDOMail(Email,Topic,TextBody)
        'dim     objCDOMail
        'Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
        'objCDOMail.From ="cs@jerbo.cn" '改为你的邮箱
        'objCDOMail.To = Email
        'objCDOMail.Subject = Topic
		'objCDOMail.BodyFormat=0
		'objCDOMail.MailFormat=0
        'objCDOMail.Body = TextBody
		
        'objCDOMail.Send
        'Set objCDOMail = Nothing
        'SendCDOMail = 1
        dim Mailsend
        set Mailsend = Server.CreateObject("easymail.Mailsend")
        Dim Tid,Un
        Un = "cs@jerbo.cn"     '您的邮件服务器登录名，不需要密码
        Dim EI
        Set EI = server.CreateObject("easymail.Users")
        Tid = EI.Login(un)
        Set EI = Nothing
        Mailsend.createnew Un,Tid '邮箱账号,临时ID
        Mailsend.CharSet = "gb2312"     '编码
        Mailsend.MailName = "Http://www.jerbo.cn(南京旅游网)"     '发件人名

        Mailsend.EM_BackAddress = "" '邮件回复地址
        Mailsend.EM_Bcc = "" '暗送地址
        Mailsend.EM_Cc = "" '抄送地址
        Mailsend.EM_OrMailName = "" '原邮件名
        Mailsend.EM_Priority = "Normal" '邮件重要度      
        Mailsend.EM_ReadBack = false '是否读取确认,挂号信(限本系统内用户)      
        Mailsend.EM_SignNo = -1     '使用签名的序号
     
        Mailsend.EM_Subject = Topic '主题
        'Mailsend.EM_Text = TextBody '内容
        Mailsend.EM_HTML_Text = TextBody 'HTML邮件内容
        Mailsend.useRichEditer = true '发送的是否为HTML格式邮件

        Mailsend.EM_TimerSend = ""     '定时发送的时间
        Mailsend.EM_To = Email '收件人地址
        Mailsend.ForwardAttString = "" '转发邮件时的原附件

        Mailsend.AddFromAttFileString = "" '添加自网络存储中的文件名

        Mailsend.SystemMessage = false '是否是系统邮件

        Mailsend.SendBackup = false '是否保存发送邮件
     
        If Mailsend.Send() = false Then
              SendEasyMail = 0
        Else
              SendEasyMail = 1
        End If
        Set Mailsend = nothing
End Function

function Delhtml(str)
    dim re
    Set re=new RegExp
    re.IgnoreCase =true
    re.Global=True
    re.Pattern="(\<.[^\<]*\>)"
    str=re.replace(str," ")
    re.Pattern="(\<\/[^\<]*\>)"
    str=re.replace(str," ")
    Delhtml=str
    set re=nothing
end function

function DecodeUrl(str)
	dim re
	Set re = New RegExp
	re.Pattern = "(http\:\/\/|^http\:\/\/)([\w-]+\.+[\w-]+\.+[\w-]+\.+[\w-]+|[\w-]+\.+[\w-]+\.+[\w-]+)(/[\w-./?%&=]*)?"
	re.Global = true
	re.IgnoreCase = true
	str = re.Replace(str,"<a href=http://$2$3 target=_blank class=blue001>$1$2$3</a>")
	DecodeUrl = str
end function

Function GetUrl()
  Dim ScriptAddress, M_ItemUrl, M_item
  ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME")) '取得当前地址
  M_ItemUrl = ""
  If (Request.QueryString <> "") Then
  ScriptAddress =  ScriptAddress & "?"
  For Each M_item In Request.QueryString
   If InStr(page,M_Item)=0 Then
    M_ItemUrl = M_ItemUrl & M_Item &"="& Server.URLEncode(Request.QueryString(""&M_Item&""))  & "&"
   End If
  Next
  end if
  GetUrl = ScriptAddress & M_ItemUrl
 End Function
 
 Function NewSendMail(Email,Topic,TextBody)
	dim    objCDOMail
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From ="cs@jerbo.cn" '改为你的邮箱
	objCDOMail.To = Email
	objCDOMail.Subject = Topic
	objCDOMail.BodyFormat=0
	objCDOMail.MailFormat=0
	objCDOMail.Body = TextBody
	
	objCDOMail.Send
	Set objCDOMail = Nothing
	SendCDOMail = 1
 End Function
'过滤脚本
Function ClearScript(Str)
	Set re = New RegExp
	re.Pattern = "(<)([script]*)(>)"'(*)(<\/)(*script)(>)
	re.Global = true
	re.IgnoreCase = true
	str = re.Replace(str,"&lt;$2&gt;")
	ClearScript=str
End Function
function IsNumeric_Request(str)
	dim values
	values=request(str)
	if not IsNumeric(values) then
		Jstop("非法操作")
	Else
		IsNumeric_Request=values
	End if
end function
function encodestr(str)
if str<>"" then
	str=trim(str)
	str=replace(str,"<","&lt;")
	str=replace(str,">","&gt;")
	str=replace(str,"'","""")
	str=replace(str,vbCrLf&vbCrlf,"</p><p>")
	str=replace(str,vbCrLf,"<br>")
	encodestr=replace(str," ","&nbsp")
end if
end function

Function Jmail(mailTo,mailTopic,mailBody,mailCharset,mailContentType)
	'--------------------------------------------------------------------
	'JMail
	'--------------------------------------------------------------------
	'入口参数：
	'　　　　mailTo 收件人email地址
	'　　　　mailTopic 邮件主题
	'　　　　mailBody 邮件正文(内容)
	'　　　　mailCharset 邮件字符集，例如GB2312或US-ASCII
	'　　　　mailContentType 邮件正文格式，例如text/plain或text/html
	'返回值：
	'　　　　字符串，发送成功后返回OK，不成功返回错误信息
	'使用方法：
	'　　　　1)设置好常量，即以Const开头的变量
	'　　　　2)使用类似如下代码发信
	'Dim SendStat
	'SendStat = Jmail("aa@163.com","测试Jmail","这是一封<br/>测试信！","GB2312","text/html")
	'Response.Write SendStat
	'--------------------------------------------------------------------
	
	'***************根据需要设置常量开始*****************
	Dim ConstFromNameCn,ConstFromNameEn,ConstFrom,ConstMailDomain,ConstMailServerUserName,ConstMailServerPassword
	
	ConstFromNameCn = "南京旅游网"'发信人中文姓名(发中文邮件的时候使用)，例如'张三'
	ConstFromNameEn = "taozhi"'发信人英文姓名(发英文邮件的时候使用)，例如'zhangsan'
	ConstFrom = "service@jerbo.cn"'发信人邮件地址，例如'zhangsan@163.com'
	ConstMailDomain = "mail.jerbo.cn"'smtp服务器地址，例如smtp.163.com
	ConstMailServerUserName = "service@jerbo.cn"'smtp服务器的信箱登陆名，例如'zhangsan'。注意要与发信人邮件地址一致！
	ConstMailServerPassword = "#)!$$#@!"'smtp服务器的信箱登陆密码
	'***************根据需要设置常量结束*****************
	
	'-----------------------------以下内容无需改动------------------------------
	On Error Resume Next
	Dim myJmail
	Set myJmail = Server.CreateObject("JMail.Message")
	myJmail.Logging = True'记录日志
	myJmail.ISOEncodeHeaders = False'邮件头不使用ISO-8859-1编码
	myJmail.ContentTransferEncoding = "base64"'邮件编码设为base64
	myJmail.AddHeader "Priority","3"'添加邮件头,不要改动！
	myJmail.AddHeader "MSMail-Priority","Normal"'添加邮件头,不要改动！
	myJmail.AddHeader "Mailer","Microsoft Outlook Express 6.00.2800.1437"'添加邮件头,不要改动！
	myJmail.AddHeader "MimeOLE","Produced By Microsoft MimeOLE V6.00.2800.1441"'添加邮件头,不要改动！
	myJmail.Charset = mailCharset
	myJmail.ContentType = mailContentType
	
	If UCase(mailCharset) = "GB2312" Then
	myJmail.FromName = ConstFromNameCn
	Else
	myJmail.FromName = ConstFromNameEn
	End If
	
	myJmail.From = ConstFrom
	myJmail.Subject = mailTopic
	myJmail.Body = mailBody
	myJmail.AddRecipient mailTo
	myJmail.MailDomain = ConstMailDomain
	myJmail.MailServerUserName = ConstMailServerUserName
	myJmail.MailServerPassword = ConstMailServerPassword
	myJmail.Send ConstMailDomain
	myJmail.Close
	Set myJmail=nothing 
	
	If Err Then 
	Jmail=Err.Description
	Err.Clear
	Else
	Jmail="OK"
	End If
	
	On Error Goto 0
End Function

'过滤javascript
function movejs(str)
	dim objregexp,str1,a
	set objregexp=new regexp	
	objregexp.ignorecase =true
	objregexp.global=true
	objregexp.pattern="\<script.+?\<\/script\>"
	a=objregexp.replace(str,"")	
	objregexp.pattern="\<[^\<]+>"
	movejs=objregexp.replace(a,"")
end function 


'过滤html标签只剩<br>
function filterhtml(byval fstring)
    if isnull(fstring) or trim(fstring)="" then
        filterhtml=""
        exit function
    end if
    fstring = replace(fstring, "<br />", "[br]")
    fstring = replace(fstring, "<br>", "[br]")
    '过滤html标签
    dim re
    set   re = new   regexp
    re.ignorecase=true
    re.global=true
    re.pattern="<(.+?)>"
    fstring = re.replace(fstring, "")
    set   re=nothing
    fstring = replace(fstring, "[br]", "<br />")
    filterhtml = fstring
end function


'********************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
'********************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'***************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'***************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function


Function CreateNewsPage(ReplaceString,ValueString,ModelPath,SaveFolderPath,TimeFolderPath,SaveFileName)
'CreateNewsPage生成新闻页面
'ReplaceString被替换标签串，ValueString替换值串，ModelPath模板路径（相对路径），SaveFolderPath生成的页面保存文件夹（绝对路径）
	Dim TempletContent,FolderName
	If Right(SaveFolderPath,1)="/" or Right(SaveFolderPath,1)="\" Then SaveFolderPath=Left(Len(SaveFolderPath)-1) 
	'ModelPath=Server.Mappath(ModelPath)
	Dim Fso
	Set Fso=Server.CreateObject("Scripting.FileSystemObject")

	If Not Fso.FileExists(ModelPath) Then
		
		CreateNewsPage="模板文件不存在，操作中断。"
		Exit Function
	End If
	If Not Fso.FolderExists(SaveFolderPath) then
		CreateNewsPage="待保存文件夹不存在，操作中断。"
		Exit Function
	end if
	
	'时间文件夹
	dim TimeFolderPath_
	if TimeFolderPath="" then TimeFolderPath=year(date)&month(date)
	TimeFolderPath_=SaveFolderPath&"/"&TimeFolderPath
	If Not Fso.FolderExists(TimeFolderPath_) then
		Fso.CreateFolder(TimeFolderPath_)
	end if
	dim ReadModel,ReStr,VaStr,Min_ReStr,Max_VaStr,i,OldFileArr,TempPath
	Set ReadModel = Fso.OpenTextFile(ModelPath,1,True)
	If ReadModel.AtEndOfStream<> true Then
		TempletContent= ReadModel.ReadAll
	End If
	ReadModel.Close()
	Set ReadModel=Nothing
	ReStr=Split(ReplaceString,",")
	VaStr=Split(ValueString,"$$$$")
	Min_ReStr=LBound(ReStr)
	Max_VaStr=UBound(VaStr)
	For i=Min_ReStr To Max_VaStr
		TempletContent=Replace(TempletContent,ReStr(i),VaStr(i))
	Next
	
	dim NewsPageWrite
	If SaveFileName<>"" Then   '如果文件名存在
		If Fso.FileExists(SaveFolderPath&"/"&TimeFolderPath&"/"&SaveFileName) Then
			Fso.DeleteFile SaveFolderPath&"/"&TimeFolderPath&"/"&SaveFileName,true
		Else
			OldFileArr=Split(SaveFileName,"/")
			TempPath=SaveFolderPath&"/"
			For i=LBound(OldFileArr) to (UBound(OldFileArr)-1)
				TempPath=TempPath&"/"&OldFileArr(i)
				If Not Fso.FolderExists(TempPath) then
					Fso.CreateFolder(TempPath)
				end if
			Next
		End If
		
		Set NewsPageWrite = Fso.CreateTextFile(SaveFolderPath&"/"&TimeFolderPath&"/"&SaveFileName)
		NewsPageWrite.WriteLine(TempletContent)
		NewsPageWrite.Close
		Set NewsPageWrite= Nothing
		set Fso=Nothing
		CreateNewsPage=TimeFolderPath&"@@@@"&SaveFileName
	Else
		Randomize
		dim RanNum,TheNewsPage
		RanNum=int(90000*rnd)+10000
		TheNewsPage=day(now)&hour(now)&minute(now)&second(now)&RanNum&".asp"
		Set NewsPageWrite = Fso.CreateTextFile(SaveFolderPath&"/"&TimeFolderPath&"/"&TheNewsPage)
		NewsPageWrite.WriteLine(TempletContent)
		NewsPageWrite.Close
		Set NewsPageWrite= Nothing
		set Fso=Nothing
		CreateNewsPage=TimeFolderPath&"@@@@"&TheNewsPage
	End If

End Function

function cleanWord(html)
    dim regEx
    set regEx=New RegExp
    regEx.IgnoreCase=True
    regEx.Global=True
    regEx.Pattern="<[^>]*>"                    '清除所有<>之间的内容
    html = regEx.replace(html,"" )
    regEx.Pattern="{[^}]*}"                     '清除所有{}之间的内容
    html = regEx.replace(html,"" )
    regEx.Pattern="/[^/]*/"                       '清除所有/**/之间的注释
    html = regEx.replace(html,"" )
    html =Replace(html,"table.MsoNormalTable","")        '替换掉漏网的单词
    cleanWord= html 
    set regEx=nothing
end function
%>
<script language="javascript">
//----------- 验证数字 ----------
function checkNum(Num){
  var Expression=/^[1-9]+[0-9]*\d$/;
  var re=new RegExp(Expression);
  if(re.test(Num)==true){
    return true;}
  else{
    return false;}
} 
//----------- 验证E-mail地址 ---------- 
function checkEmail(email){
  var Expression=/\w+([-+.']\w+)*\.\w+([-.]\w+)*/;
  var re=new RegExp(Expression);
  if(re.test(email)==true){
    return true;}
  else{
    return false;}
}
//----------- 验证网址 ---------- 
function checkUrl(url){
  var Expression=/http(s)?:\/\/([\w-]+\.)+[\w-]+(\/[\w-.\/?%&=]*)?/;
  var re=new RegExp(Expression);
  if(re.test(url)==true){
    return true;}
  else{
    return false;}
}
//----------- 验证身份证号码 ----------
function checkCode(code) {
  var Expression=/\d{17}[\d|X]|\d{15}/;
  var re=new RegExp(Expression);
  if(re.test(code)==true){
    return true;}
  else{
    return false;}
}
</script>
