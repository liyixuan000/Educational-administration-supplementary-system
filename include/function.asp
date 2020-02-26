<%
'---------- ����ת���ַ����� -----------
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
'---------- ����ʱ���ȡ���ַ��� -----------
Function GetOrderNo(dDate)
    GetOrderNo = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)+RIGHT("00" + Trim(Hour(dDate)),2)+RIGHT("00"+Trim(Minute(dDate)),2)+RIGHT("00"+Trim(Second(dDate)),2)
End Function
'---------- ��ȡ��������� -----------
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
'aΪҪȥ�ַ�������,bΪ��ʼ�ַ�,cΪ�����ַ�   ȥ����ĳ���ַ�����һ�ַ�֮����ַ���  
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
   '���������Ƿ�������
   'str ������ַ�����def���str�Ƿ��򷵻ص�����
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
'��������GetStrLen
'��  �ã�'ȡ�ַ������ȣ�����Ϊ�����ֽ�
'��  ����str   ----ԭ�ַ���
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
'�������ã�������֤�����Ƿ���ȷ
'=============================================================
Function CheckCardId(e)   
    arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")   
    Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")   
    Checker = Split("1,9,8,7,6,5,4,3,2,1,1", ",")   
    If Len(e) < 15 Or Len(e) = 16 Or Len(e) = 17 Or Len(e) > 18 Then   
        CheckCardId= "���֤�Ź��� 15 ���18λ"     
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
        CheckCardId= "���֤�����һλ�⣬����Ϊ���֣�"           
        Exit Function   
    End If   
    Dim strYear, strMonth, strDay   
    strYear = CInt(Mid(Ai, 7, 4))   
    strMonth = CInt(Mid(Ai, 11, 2))   
    strDay = CInt(Mid(Ai, 13, 2))   
    BirthDay = Trim(strYear) + "-" + Trim(strMonth) + "-" + Trim(strDay)   
    If IsDate(BirthDay) Then   
        If DateDiff("yyyy",Now,BirthDay)<-140 or cdate(BirthDay)>date() Then           
            CheckCardId= "���֤�������"   
            Exit Function   
        End If   
        If strMonth > 12 Or strDay > 31 Then   
            CheckCardId= "���֤�������"   
            Exit Function   
        End If   
    Else   
        CheckCardId= "���֤�������"   
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
		CheckCardId= "���֤�������"   
        Exit Function   
    End If 
	CheckCardId ="True"  
End Function 

'=============================================================
'�������ã�ƥ��������ʽ
'=============================================================
Function CheckExp(patrn, strng) 
	Dim regEx,Match'���������� 
	Set regEx = New RegExp'����������ʽ�� 
	regEx.Pattern = patrn'����ģʽ�� 
	regEx.IgnoreCase = true'�����Ƿ������ַ���Сд�� 
	regEx.Global = True'����ȫ�ֿ����ԡ� 
	Matches = regEx.test(strng)'ִ�������� 
	CheckExp = matches
End Function

'-------------------------------------------------------------
' Function Name : IsValidString
' Function Desc : �ж������Ƿ���һ���� 0-9 / A-Z / a-z / . / _ ��ɵ��ַ���
'-------------------------------------------------------------
Function IsValidString(sInput)
	
	Dim oRegExp
	
	'����������ʽ
	Set oRegExp = New RegExp
	
	'����ģʽ
	oRegExp.Pattern = "^[a-zA-Z0-9\s.\-_]+$"
	'�����Ƿ������ַ���Сд
	oRegExp.IgnoreCase = True
	'����ȫ�ֿ�����
	oRegExp.Global = True
	
	'ִ������
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

'��ȫ����
Function SafeDiv(A,B)
	SafeDiv = checknumber(A,0) / checknumber(B,1)
End Function

function checksqlstr(getstr)
'�������Ĳ����Ƿ���sql�����ַ�������з��ؿ��ַ���
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
	
	
	regex.Pattern="��[\s��]*��[\s��]*��"
	getstr=regex.Replace(getstr,"*��*")
	set regex=nothing
	checksqlstr = getstr
end function

function checkuserstr(str)
   '����û�ע����
	if len(str) = 0 or isnull(str) then
		checkuserstr = ""
		exit function
	end if
	dim i 
	dim lb_user,strfilter,regEx,str1
	str1 = "[A-Za-z0-9|\u4e00-\u9fa5]*" '��������ݱ���������
	strfilter="\$|\(|\)|\*|\+|\-|\.|\[|]|\?|\\|\^|\{|\||}|~|`|!|@|#|%|&|_|=|<|>|/|,|'| |��|\:|cubcn|admin|administrator|administrators|luo|����|system"
	Set regEx = New RegExp
	regEx.Pattern =	strfilter
	regEx.IgnoreCase = True
	lb_user = regex.Test(str)
	if not lb_user then
		regex.Pattern="��[ ��]*��[ ��]*��"
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
	'strfilter1 = "^[A-Za-z0-9\u4e00-\u9fa5]+$" '��������ݱ���������
	'Set regEx2 = New RegExp
	'regEx2.Pattern =	strfilter1
	'regEx2.IgnoreCase = True
	
	'if not regex2.Test(str) then		
	'	checkuserstr = ""
	'end if
	'set regex2 = nothing
	
	set regex=nothing	
end function 



'����ļ�·��
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

'������·��,ת־��ǰ����·��
Function PathToCurrent(FpathStr)
	PathToCurrent = Server.MapPath(FpathStr)
End Function


'�����ļ����ĺ�׺ �������ļ���
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
'��ʾ��Ϣ������һ��
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


Rem �õ�����
Function GetTxt(Path)
	On Error resume next
	Dim Ph,fso,ts
	Set fso = CreateObject("Scripting.FileSystemObject")
	ph=server.mappath(Path)
	Set ts = fso.OpenTextFile(ph, 1)
	GetTxt = ts.ReadAll
	ts.Close
End Function

Rem ���Ŀ¼�Ƿ����,����������Ŀ¼
Function CreateFolder(Folder)
	DIM FSO
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if Not fso.FolderExists(Server.MapPath(Folder)) Then 
		fso.CreateFolder(Server.MapPath(Folder))
	End if
	Set Fso = Nothing
	If Err.number<>0 Then
	 err.Clear
	  Response.Write "Ȩ�޴���!!��������IISĬ�ϵ������û��Ƿ�Ա�ϵͳ���ڵ�Ŀ¼�ж�,д,�޸�Ȩ��!!2"
	  Response.End
	End If
End Function
Rem ɾ���ļ�
Function DelFile(Path)
	DIM fso
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	IF fso.FileExists(Server.MapPath(Path)) then
		fso.DeleteFile(Server.MapPath(Path))
	End IF
	Set fso=Nothing
End Function


Rem ɾ��Ŀ¼
Function DelFolder(Path)
	DIM fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	IF fso.FolderExists(Server.MapPath(Path)) then
		fso.DeleteFolder(Server.MapPath(Path))
	End IF
	Set fso=Nothing
End Function


Rem �����ļ�
Sub copyfile(movefiles, copyfiles)
	movefiles = Server.MapPath(movefiles)
	copyfiles = Server.MapPath(copyfiles)
	Dim Fso
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fso.CopyFile movefiles, copyfiles

	Set fso = Nothing
End Sub

Rem ����û��ϴ�Ŀ¼�Ƿ�����
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


REM д���ļ�
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
  fname = replace(fname,"����","")
  fname = replace(fname,"����","")
  makefilename=fname
end Function

'===================================================================
'                           ʱ��ת������2
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
'                           ʱ��ת������3
'===================================================================
function chan_day(shijian)	
	Dim s_day
	s_day=day(shijian)	
	chan_day=s_day 
	
end function

'===================================================================
'�������ܣ��滻��������ģ������
'����˵����(��)
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
	TempletContent=replace(TempletContent,"$Join_Record$",Join_Record)		'��ؼ�¼
End Function	

'***********************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
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
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
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
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table></form>"
	response.write strTemp
End Sub

'***********************************************
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
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
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	response.write strTemp
End Sub

'***********************************************
'��������Show_Page
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
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
'��������infotime
'��  �ã���ʽ������
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
'��������infotime_F,���ַ�
'��  �ã���ʽ������
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
		Response.Write("�벻Ҫ��������վ�ύ���ݵ���վ������")
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
		'���Ӷ������ַ��ı���
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


'ȡ��ҳ����
Function GetHttp(url) 
 Dim Retrieval,GetBody
  Set Retrieval = Server.CreateObject("MSXML2.serverXMLHTTP") '//�ѵ��ʲ𿪷�ֹɱ�������ɱ
  Retrieval.Open "GET",url,false
  Retrieval.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
  Retrieval.Send
  GetBody=Retrieval.ResponseBody 
  GetHttp=BytesToBstr(GetBody)
  Set Retrieval = Nothing 
End function


'������ת�ı�
Function BytesToBstr(body)
  Dim objstream
   Set objstream = CreateObject("Ado" & "db.Str" & "eam") '//�ѵ��ʲ𿪷�ֹɱ�������ɱ
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

'��CDOnts������ʼ�
Function SendCDOMail(Email,Topic,TextBody)
        'dim     objCDOMail
        'Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
        'objCDOMail.From ="cs@jerbo.cn" '��Ϊ�������
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
        Un = "cs@jerbo.cn"     '�����ʼ���������¼��������Ҫ����
        Dim EI
        Set EI = server.CreateObject("easymail.Users")
        Tid = EI.Login(un)
        Set EI = Nothing
        Mailsend.createnew Un,Tid '�����˺�,��ʱID
        Mailsend.CharSet = "gb2312"     '����
        Mailsend.MailName = "Http://www.jerbo.cn(�Ͼ�������)"     '��������

        Mailsend.EM_BackAddress = "" '�ʼ��ظ���ַ
        Mailsend.EM_Bcc = "" '���͵�ַ
        Mailsend.EM_Cc = "" '���͵�ַ
        Mailsend.EM_OrMailName = "" 'ԭ�ʼ���
        Mailsend.EM_Priority = "Normal" '�ʼ���Ҫ��      
        Mailsend.EM_ReadBack = false '�Ƿ��ȡȷ��,�Һ���(�ޱ�ϵͳ���û�)      
        Mailsend.EM_SignNo = -1     'ʹ��ǩ�������
     
        Mailsend.EM_Subject = Topic '����
        'Mailsend.EM_Text = TextBody '����
        Mailsend.EM_HTML_Text = TextBody 'HTML�ʼ�����
        Mailsend.useRichEditer = true '���͵��Ƿ�ΪHTML��ʽ�ʼ�

        Mailsend.EM_TimerSend = ""     '��ʱ���͵�ʱ��
        Mailsend.EM_To = Email '�ռ��˵�ַ
        Mailsend.ForwardAttString = "" 'ת���ʼ�ʱ��ԭ����

        Mailsend.AddFromAttFileString = "" '���������洢�е��ļ���

        Mailsend.SystemMessage = false '�Ƿ���ϵͳ�ʼ�

        Mailsend.SendBackup = false '�Ƿ񱣴淢���ʼ�
     
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
  ScriptAddress = CStr(Request.ServerVariables("SCRIPT_NAME")) 'ȡ�õ�ǰ��ַ
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
	objCDOMail.From ="cs@jerbo.cn" '��Ϊ�������
	objCDOMail.To = Email
	objCDOMail.Subject = Topic
	objCDOMail.BodyFormat=0
	objCDOMail.MailFormat=0
	objCDOMail.Body = TextBody
	
	objCDOMail.Send
	Set objCDOMail = Nothing
	SendCDOMail = 1
 End Function
'���˽ű�
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
		Jstop("�Ƿ�����")
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
	'��ڲ�����
	'��������mailTo �ռ���email��ַ
	'��������mailTopic �ʼ�����
	'��������mailBody �ʼ�����(����)
	'��������mailCharset �ʼ��ַ���������GB2312��US-ASCII
	'��������mailContentType �ʼ����ĸ�ʽ������text/plain��text/html
	'����ֵ��
	'���������ַ��������ͳɹ��󷵻�OK�����ɹ����ش�����Ϣ
	'ʹ�÷�����
	'��������1)���úó���������Const��ͷ�ı���
	'��������2)ʹ���������´��뷢��
	'Dim SendStat
	'SendStat = Jmail("aa@163.com","����Jmail","����һ��<br/>�����ţ�","GB2312","text/html")
	'Response.Write SendStat
	'--------------------------------------------------------------------
	
	'***************������Ҫ���ó�����ʼ*****************
	Dim ConstFromNameCn,ConstFromNameEn,ConstFrom,ConstMailDomain,ConstMailServerUserName,ConstMailServerPassword
	
	ConstFromNameCn = "�Ͼ�������"'��������������(�������ʼ���ʱ��ʹ��)������'����'
	ConstFromNameEn = "taozhi"'������Ӣ������(��Ӣ���ʼ���ʱ��ʹ��)������'zhangsan'
	ConstFrom = "service@jerbo.cn"'�������ʼ���ַ������'zhangsan@163.com'
	ConstMailDomain = "mail.jerbo.cn"'smtp��������ַ������smtp.163.com
	ConstMailServerUserName = "service@jerbo.cn"'smtp�������������½��������'zhangsan'��ע��Ҫ�뷢�����ʼ���ַһ�£�
	ConstMailServerPassword = "#)!$$#@!"'smtp�������������½����
	'***************������Ҫ���ó�������*****************
	
	'-----------------------------������������Ķ�------------------------------
	On Error Resume Next
	Dim myJmail
	Set myJmail = Server.CreateObject("JMail.Message")
	myJmail.Logging = True'��¼��־
	myJmail.ISOEncodeHeaders = False'�ʼ�ͷ��ʹ��ISO-8859-1����
	myJmail.ContentTransferEncoding = "base64"'�ʼ�������Ϊbase64
	myJmail.AddHeader "Priority","3"'����ʼ�ͷ,��Ҫ�Ķ���
	myJmail.AddHeader "MSMail-Priority","Normal"'����ʼ�ͷ,��Ҫ�Ķ���
	myJmail.AddHeader "Mailer","Microsoft Outlook Express 6.00.2800.1437"'����ʼ�ͷ,��Ҫ�Ķ���
	myJmail.AddHeader "MimeOLE","Produced By Microsoft MimeOLE V6.00.2800.1441"'����ʼ�ͷ,��Ҫ�Ķ���
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

'����javascript
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


'����html��ǩֻʣ<br>
function filterhtml(byval fstring)
    if isnull(fstring) or trim(fstring)="" then
        filterhtml=""
        exit function
    end if
    fstring = replace(fstring, "<br />", "[br]")
    fstring = replace(fstring, "<br>", "[br]")
    '����html��ǩ
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
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
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
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
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
'CreateNewsPage��������ҳ��
'ReplaceString���滻��ǩ����ValueString�滻ֵ����ModelPathģ��·�������·������SaveFolderPath���ɵ�ҳ�汣���ļ��У�����·����
	Dim TempletContent,FolderName
	If Right(SaveFolderPath,1)="/" or Right(SaveFolderPath,1)="\" Then SaveFolderPath=Left(Len(SaveFolderPath)-1) 
	'ModelPath=Server.Mappath(ModelPath)
	Dim Fso
	Set Fso=Server.CreateObject("Scripting.FileSystemObject")

	If Not Fso.FileExists(ModelPath) Then
		
		CreateNewsPage="ģ���ļ������ڣ������жϡ�"
		Exit Function
	End If
	If Not Fso.FolderExists(SaveFolderPath) then
		CreateNewsPage="�������ļ��в����ڣ������жϡ�"
		Exit Function
	end if
	
	'ʱ���ļ���
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
	If SaveFileName<>"" Then   '����ļ�������
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
    regEx.Pattern="<[^>]*>"                    '�������<>֮�������
    html = regEx.replace(html,"" )
    regEx.Pattern="{[^}]*}"                     '�������{}֮�������
    html = regEx.replace(html,"" )
    regEx.Pattern="/[^/]*/"                       '�������/**/֮���ע��
    html = regEx.replace(html,"" )
    html =Replace(html,"table.MsoNormalTable","")        '�滻��©���ĵ���
    cleanWord= html 
    set regEx=nothing
end function
%>
<script language="javascript">
//----------- ��֤���� ----------
function checkNum(Num){
  var Expression=/^[1-9]+[0-9]*\d$/;
  var re=new RegExp(Expression);
  if(re.test(Num)==true){
    return true;}
  else{
    return false;}
} 
//----------- ��֤E-mail��ַ ---------- 
function checkEmail(email){
  var Expression=/\w+([-+.']\w+)*\.\w+([-.]\w+)*/;
  var re=new RegExp(Expression);
  if(re.test(email)==true){
    return true;}
  else{
    return false;}
}
//----------- ��֤��ַ ---------- 
function checkUrl(url){
  var Expression=/http(s)?:\/\/([\w-]+\.)+[\w-]+(\/[\w-.\/?%&=]*)?/;
  var re=new RegExp(Expression);
  if(re.test(url)==true){
    return true;}
  else{
    return false;}
}
//----------- ��֤���֤���� ----------
function checkCode(code) {
  var Expression=/\d{17}[\d|X]|\d{15}/;
  var re=new RegExp(Expression);
  if(re.test(code)==true){
    return true;}
  else{
    return false;}
}
</script>
