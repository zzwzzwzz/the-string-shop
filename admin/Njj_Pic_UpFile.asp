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

Sfileup="300��jpg,bmp,gif,png,jpeg"
Sfileups=split(Sfileup,"��")
filesize=Sfileups(0)

Server.ScriptTimeOut=5000
response.Buffer=true
FormPath="../uploadpic/"
call Upload_0()
Sub Upload_0()
	Set Upload = New UpFile_Class						''�����ϴ�����
	Upload.InceptFileType = Fileli(Sfileups(1))		'�ϴ���������
	Upload.MaxSize = Int(filesize)*1024	'���ƴ�С
	Upload.GetDate()	'ȡ���ϴ�����
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Case 2 : Response.Write "ͼƬ��С����������"&filesize&"k��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Case 3 : Response.Write "���ϴ����Ͳ���ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		End Select
		Exit Sub
	Else
			Fname=Upload.Form("Fname")
	        flag=Upload.Form("flag")
		 For Each FormName in Upload.file		''�г������ϴ��˵��ļ�
			 Set File = Upload.File(FormName)	''����һ���ļ�����
			 If File.Filesize<10 Then
		 		Response.Write "����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				Exit Sub
	 		End If
			FileExt	= FixName(File.FileExt)
 			If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 				Response.Write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				Exit Sub
			End If
			Filenamelong=UserFaceName(FileExt)
 			FileName=FormPath&Filenamelong
 		
 			
 			
 			If File.FileSize>0 Then   ''��� FileSize > 0 ˵�����ļ�����
				File.SaveToFile Server.mappath(FileName)   ''�����ļ�
				
 if flag="c2" then
 response.write "<center><FIELDSET align=center><LEGEND align=center><font color=red>�ļ��ϴ��ɹ� </font></LEGEND><br>[ <a href=# onclick=""Addpic('"&d&"uploadpic/"&Filenamelong&"')"">����������ӵ��༭����</a> ]</fieldset>"
 response.end
 else
'�ͷ��ϴ�����
 response.write "<script>opener.document.form1."&Fname&".value='"&Filenamelong&"'</script>"
 response.write "<script>alert(""ͼƬ�ϴ��ɹ�!"");window.close();</script>" 
 response.end
 response.Write("<br>��������û��������,�ļ�������ȷ,û���쳣,�ļ��ϴ��ɹ�.")
 end if
				Response.Write "ͼƬ�ϴ��ɹ�!"
 			End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
End Sub


'�ж��ļ������Ƿ�ϸ�
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
'��ʽ��׺
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
'�ļ�Content-Type�ж�
Private Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
End Function

'�ļ�������
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