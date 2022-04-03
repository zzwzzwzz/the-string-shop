<!--#include file="admin_check.asp"-->
<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Upload.inc"-->
<link rel="stylesheet"  href="style.css" type="text/css">
<script  language="JavaScript">
<!-- Hide from older browsers...
//Function to add pic
function Addpic(imagePath){	
	window.opener.frames.message.focus();								
	window.opener.frames.message.document.execCommand('InsertImage', false, imagePath);
window.close();
}
// -->
</script>
<%
dim url,a,b,c,d,x,url1,e
url=Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("PATH_INFO")
dim dirnow
dirnow=split(url,"/")
a=dirnow(ubound(dirnow))
e=dirnow((ubound(dirnow)-1))
url1="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("PATH_INFO")
c=len(url1)-len(a)-len(e)-1
d=left(url1,c) 

Sfileup="300‖jpg,bmp,gif,png,jpeg"
Sfileups=split(Sfileup,"‖")
filesize=Sfileups(0)

Server.ScriptTimeOut=5000
response.Buffer=true
FormPath="../uploadpic/"
call Upload_0()
Sub Upload_0()
	Set Upload = New UpFile_Class						''建立上传对象
	Upload.InceptFileType = Fileli(Sfileups(1))		'上传类型限制
	Upload.MaxSize = Int(filesize)*1024	'限制大小
	Upload.GetDate()	'取得上传数据
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Case 2 : Response.Write "图片大小超过了限制"&filesize&"k　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Case 3 : Response.Write "所上传类型不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		End Select
		Exit Sub
	Else
			Fname=Upload.Form("Fname")
	        flag=Upload.Form("flag")
		 For Each FormName in Upload.file		''列出所有上传了的文件
			 Set File = Upload.File(FormName)	''生成一个文件对象
			 If File.Filesize<10 Then
		 		Response.Write "请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				Exit Sub
	 		End If
			FileExt	= FixName(File.FileExt)
 			If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 				Response.Write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				Exit Sub
			End If
			Filenamelong=UserFaceName(FileExt)
 			FileName=FormPath&Filenamelong
 		
 			
 			
 			If File.FileSize>0 Then   ''如果 FileSize > 0 说明有文件数据
				File.SaveToFile Server.mappath(FileName)   ''保存文件
				
 if flag="c2" then
 response.write "<center><FIELDSET align=center><LEGEND align=center><font color=red>文件上传成功 </font></LEGEND><br>[ <a href=# onclick=""Addpic('"&d&"uploadpic/"&Filenamelong&"')"">点击这里添加到编辑器中</a> ]</fieldset>"
 response.end
 else
'释放上传对象
 response.write "<script>opener.document.form1."&Fname&".value='"&Filenamelong&"'</script>"
 response.write "<script>alert(""图片上传成功!"");window.close();</script>" 
 response.end
 response.Write("<br>总数据量没超过限制,文件类型正确,没有异常,文件上传成功.")
 end if
				Response.Write "图片上传成功!"
 			End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
End Sub


'判断文件类型是否合格
Private Function CheckFileExt(FileExt)
	Dim ForumUpload,i
	ForumUpload=Fileli(Sfileups(1))
	ForumUpload=Split(ForumUpload,",")
	CheckFileExt=False
	For i=0 to UBound(ForumUpload)
		If LCase(FileExt)=Lcase(Trim(ForumUpload(i))) Then
			CheckFileExt=True
			Exit Function
		End If
	Next
End Function
'格式后缀
Function FixName(UpFileExt)
	If IsEmpty(UpFileExt) Then Exit Function
	FixName = Lcase(UpFileExt)
	FixName = Replace(FixName,Chr(0),"")
	FixName = Replace(FixName,".","")
	FixName = Replace(FixName,"asp","")
	FixName = Replace(FixName,"asa","")
	FixName = Replace(FixName,"aspx","")
	FixName = Replace(FixName,"cer","")
	FixName = Replace(FixName,"cdx","")
	FixName = Replace(FixName,"htr","")
End Function
'文件Content-Type判断
Private Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
End Function

'文件名明名
Private Function UserFaceName(FileExt)
	Dim UserID,RanNum
	'UserID = ""
	'If Dvbbs.UserID>0 Then UserID = Dvbbs.UserID&"_"
	Randomize
	RanNum = Int(90000*rnd)+10000
 	UserFaceName = UserID&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&RanNum&"."&FileExt
End Function

Function Fileli(b)
a=split(b,",")
Fileli=""
for i=0 to UBound(a)
c=trim(a(i))
Fileli=Fileli&","&c
next
End Function
%>


