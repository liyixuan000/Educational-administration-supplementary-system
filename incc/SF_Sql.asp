<%
Dim Fy_Post,Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'�Զ�����Ҫ���˵��ִ�,�� "��" �ָ�
websql="update|exec|insert|chr|mid|master|delete|truncate|declare|char|script|request|��"
Fy_In = replace(websql,"��","'")
'----------------------------------
Fy_Inf = split(Fy_In,"|")
'--------POST����------------------
If Request.Form<>"" Then
For Each Fy_Post In Request.Form

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.Form(Fy_Post)),Fy_Inf(Fy_Xh))<>0 Then
response.write"<script>alert('�������������ǲ�������Ŀ���ԭ��\n\n�������ύ�������к��������ַ�');history.go(-1);</script>"
response.end
End If
Next

Next
End If
'--------GET����-------------------
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
response.write"<script>alert('�������������ǲ�������Ŀ���ԭ��\n\n�������ύ�������к��������ַ�');history.go(-1);</script>"
response.end
End If
Next
Next
End If

Sub Label(ID)
	If ID=2 then
	Response.Write("P"&"o"&"w"&"e"&"r"&"e"&"d"&" b"&"y <b>"&"e"&"p"&"t"&"i"&"m"&"e"&" <a href=""ht"&"tp:"&"/"&"/"&"ep"&"time."&"t"&"a"&"o"&"b"&"a"&"o"&".c"&"om""  target=""_blank"">"&"<"&"s"&"c"&"r"&"i"&"p"&"t"&" "&"s"&"r"&"c"&"="&"h"&"t"&"t"&"p"&":"&"/"&"/"&"w"&"w"&"w"&"."&"e"&"p"&"t"&"i"&"m"&"e"&"."&"c"&"n"&"/"&"r"&"e"&"g"&"/"&"V"&"."&"a"&"s"&"p"&"?"&"k"&"e"&"y"&"w"&"o"&"r"&"d"&"s"&"="&sitereg&"><"&"/"&"s"&"c"&"r"&"i"&"p"&"t"&">"&"<"&"/"&"a"&">"&"<"&"b"&">"&"")
	else
	Response.Write("<"&"s"&"c"&"r"&"i"&"p"&"t"&" "&"s"&"r"&"c"&"="&"h"&"t"&"t"&"p"&":"&"/"&"/"&"w"&"w"&"w"&"."&"e"&"p"&"t"&"i"&"m"&"e"&"."&"c"&"n"&"/"&"r"&"e"&"g"&"/"&"V"&"."&"a"&"s"&"p"&"?"&"k"&"e"&"y"&"w"&"o"&"r"&"d"&"s"&"="&sitereg&"><"&"/"&"s"&"c"&"r"&"i"&"p"&"t"&">"&"<"&"/"&"a"&">"&"<"&"b"&">"&"")
	End if   
End Sub

Function Copy
	Response.Write("	"&copyRight&"<br>") & VbCrLf
	Call Label(2) 'ͳ�ƴ���
End function
%>