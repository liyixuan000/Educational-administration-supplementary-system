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

%>