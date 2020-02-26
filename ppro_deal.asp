<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
'获得添加装备信息
Response.Buffer=true
formsize=Request.TotalBytes
formdata=Request.BinaryRead(formsize)


crlf=chrB(13)&chrB(10)
strflag=leftb(formdata,clng(instrb(formdata,crlf))-1)

'信息添加到数据库
Sql_2 = "Select * from tb_Hon"
Set rs_2 = Server.CreateObject("ADODB.Recordset")
rs_2.open Sql_2,conn,1,3
rs_2.AddNew

'获得表单所有元素的值
k = 1     '注意不要用i了
While instrb(k,formdata,strflag) < instrb((instrb(k,formdata,strflag)+lenb(strflag)),formdata,strflag)
	start = instrb(k,formdata,strflag) + lenb(strflag) + 2
	endsize = instrb((instrb(k,formdata,strflag)+lenb(strflag)),formdata,strflag) - start - 2
	bin_content = midb(formdata,start,endsize)
	pos1_name = instrb(bin_content,toByte("name="""))
	pos2_name = instrb(pos1_name+6,bin_content,toByte(""""))
	nametag = midb(bin_content,pos1_name+6,pos2_name-pos1_name-6)
		
	'response.write(toStr(nametag)&"<br>")
	pos1_filename = instrb(pos2_name,bin_content,toByte("filename="""))
		
		
If(pos1_filename = 0)Then
namevalue = toStr(midb(bin_content,pos2_name+5,lenb(bin_content)-pos2_name-4))
If(InStr(toStr(nametag),"EqName") > 0)Then
rs_2("ProjectID") = namevalue
'Session("UserName") = namevalue
End If
If(InStr(toStr(nametag),"EqHonor") > 0)Then
rs_2("HonorID") = namevalue
End If
If(InStr(toStr(nametag),"EqPrice") > 0)Then
rs_2("Reason") = namevalue
End If
'If(InStr(toStr(nametag),"upload") > 0)Then
'rs_2("HonTestify") = namevalue
'End If

							
						End If
	k = instrb((instrb(k,formdata,strflag)+lenb(strflag)),formdata,strflag) 
Wend
rs_2.Update
rs_2.Close
Set rs_2 = Nothing
'Response.Redirect("index.asp")%>
<script>
					alert("上传成功！");
					history.go(-1);
					</script>
<%
Response.End()
'字符串转换成二进制数
 Private function toByte(Str)
   dim i,iCode,c,iLow,iHigh
   toByte=""
   For i=1 To Len(Str)
   c=mid(Str,i,1)
   iCode =Asc(c)
   If iCode<0 Then iCode = iCode + 65535
   If iCode>255 Then
     iLow = Left(Hex(Asc(c)),2)
     iHigh =Right(Hex(Asc(c)),2)
     toByte = toByte & chrB("&H"&iLow) & chrB("&H"&iHigh)
   Else
     toByte = toByte & chrB(AscB(c))
   End If
   Next
 End function
'二进制转达换成字符串
 Private function toStr(Byt)
 	toStr=""
 	for i=1 to lenb(byt)
	blow = midb(byt,i,1)
	if  ascb(blow)>127 then	
	toStr = toStr&chr(ascw(midb(byt,i+1,1)&blow))  '这里浪费了挺多的时间。
	i = i+1
	else
	toStr = toStr&chr(ascb(blow))
	end if
	Next
 End function

%>