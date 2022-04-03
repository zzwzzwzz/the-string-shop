<!--#include file="admin_check.asp"-->
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<!--#include file="upload2.inc"-->
<%
Sfileup="300‖jpg,bmp,gif,png,jpeg"
Sfileups=split(Sfileup,"‖")

set upload=new upload_5xSoft
formPath=upload.form("filepath")
if right(formPath,1)<>"/" then formPath=formPath&"/" 
cid				 = upload.form("cid")
prod_info_name   = upload.form("prod_info_name")
prod_info_no   	 = upload.form("prod_info_no")
prod_info_AdWord = upload.form("prod_info_AdWord")
prod_info_flag   = upload.form("prod_info_flag")
prod_info_PriceM = upload.form("prod_info_PriceM")
prod_info_PriceS = upload.form("prod_info_PriceS")
prod_info_Detail = upload.form("content")
prod_info_OnOff  = upload.form("prod_info_OnOff")

'验证商品名称重复情况
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from prod_info where prod_info_name='"&prod_info_name&"'"
rs.open sql,conn,1,3
if not(rs.eof and rs.bof) then
    response.write "<script language='javascript'>"
    response.write "alert('出错了，商品标题重复，请重新录入！');"
    response.write "location.href='javascript:history.go(-1)';"
    response.write "</script>"
    response.end
end if
rs.close
set rs=nothing

'接收图片
set file=upload.file("file") '生成一个文件对象
if file.Filesize>0 then '如果 Filesize > 0 说明有文件数据

    if file.Filesize>1024*1024 then
      	response.write "文件大小不能超过521k"
      	response.end
    end if
    
    FileExt	= FixName(File.FileExt)
			
 	If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 		Response.Write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
 		response.end
	End If

    if TrueStr(file.filename)=false then
		response.write "非法文件"
		response.end
	end if
	
    file.SaveAs Server.mappath(formPath&file.FileName)
    prod_info_picB=file.FileName
    
end if
  	  
'判断自动生成小图片还是手动上传小图片
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select top 1 root_option_PicSType from root_option"
rs.open sql,conn,1,1
root_option_PicSType=rs(0)
rs.close
set rs=nothing
if root_option_PicSType=1 then

	'接收小图片
	set file2=upload.file("file2") '生成一个文件对象
	if file2.Filesize>0 then '如果 Filesize > 0 说明有文件数据

    	if file2.Filesize>Int(filesize)*512 then
      		response.write "文件大小不能超过150kb"
      		response.end
    	end if
    
    	FileExt2	= FixName(File2.FileExt)
				
 		If Not ( CheckFileExt(FileExt2) and CheckFileType(File2.FileType) ) Then
 			Response.Write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
 			response.end
		End If

    	if TrueStr(file2.filename)=false then
			response.write "非法文件"
			response.end
		end if
	
    	file2.SaveAs Server.mappath(formPath&file2.FileName)
    	prod_info_picS=file2.FileName
    
	end if
else
  
	'=====缩略图处理========(按原图比例)
	Set Jpeg = Server.createObject("Persits.Jpeg")	'建立图片对象
	JpegPath = Server.MapPath(formPath&file.FileName)	'图片位置
	Jpeg.Open JpegPath	'打开图片
	Jpeg.Width = Jpeg.OriginalWidth
	Jpeg.Height = Jpeg.OriginalHeight
	max_width = Root_Jpeg_SPicWidth '指定缩略图最大宽
	max_height = Root_Jpeg_SPicHeight '指定缩略图最大高
	if max_width=0 or isnull(max_width)=true then max_width=80
	if max_height=0 or isnull(max_height)=true then max_height=80
	if Jpeg.OriginalWidth > max_width then
		Jpeg.Width = max_width
		Jpeg.Height = Jpeg.OriginalHeight * (Jpeg.Width / Jpeg.OriginalWidth)
		if Jpeg.Height > max_height then
			Jpeg.Height = max_height
			Jpeg.Width = Jpeg.OriginalWidth * (Jpeg.Height / Jpeg.OriginalHeight)
		end if
	else
		if Jpeg.OriginalHeight>max_height then
			Jpeg.Height=max_height
			Jpeg.Width =Jpeg.OriginalWidth * (Jpeg.Height / Jpeg.OriginalHeight)
		end if
	end if
	prod_info_picS="S"&file.FileName
	Jpeg.Save Server.MapPath(formPath&prod_info_picS)	'OK,保存缩略图文件
	Set Jpeg=nothing	'注销对象
	
end if

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from prod_info"
rs.open sql,conn,1,3
rs.addnew
rs("cid")			   = cid
rs("prod_info_name")   = prod_info_name
rs("prod_info_no")     = prod_info_no
rs("prod_info_AdWord") = prod_info_AdWord
rs("prod_info_flag")   = prod_info_flag
rs("prod_info_PriceM") = prod_info_PriceM
rs("prod_info_PriceS") = prod_info_PriceS
rs("prod_info_PicB")   = prod_info_PicB
rs("prod_info_PicS")   = prod_info_PicS
rs("prod_info_Detail") = prod_info_Detail
rs("prod_info_OnOff")  = prod_info_OnOff
rs("addtime")		   = now()
rs.update
rs.close
set rs=nothing  
call ok("您已成功添加了一条商品信息！","prod_info_add.asp?cid="&cid&"")


'判断文件类型是否合格
Function CheckFileExt(FileExt)
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
Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
End Function

Function Fileli(b)
	a=split(b,",")
	Fileli=""
	for i=0 to UBound(a)
		c=trim(a(i))
		Fileli=Fileli&","&c
	next
End Function


'******************************************************************
'ASP上传漏洞还利用"\0"对filepath进行手脚操作
'针对这样的情况可使用如下函数:
'******************************************************************
function TrueStr(fileTrue)
	str_len=len(fileTrue)
	pos=Instr(fileTrue,chr(0))
	if pos=0 or pos=str_len then
		TrueStr=true
	else
		TrueStr=false
	end if
end function
%>