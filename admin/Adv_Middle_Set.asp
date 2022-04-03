<!--#include file="admin_check.asp"-->
<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from adv_Middle where adv_Middle_id=1"
rs.open sql,conn,1,1
adv_middle_Pic1=rs("adv_middle_Pic1")
adv_middle_Pic2=rs("adv_middle_Pic2")
adv_middle_Pic3=rs("adv_middle_Pic3")
adv_middle_Pic4=rs("adv_middle_Pic4")
adv_middle_Pic5=rs("adv_middle_Pic5")
adv_middle_Pic1url=rs("adv_middle_Pic1url")
adv_middle_Pic2url=rs("adv_middle_Pic2url")
adv_middle_Pic3url=rs("adv_middle_Pic3url")
adv_middle_Pic4url=rs("adv_middle_Pic4url")
adv_middle_Pic5url=rs("adv_middle_Pic5url")
adv_middle_Pic1Txt=rs("adv_middle_Pic1Txt")
adv_middle_Pic2Txt=rs("adv_middle_Pic2Txt")
adv_middle_Pic3Txt=rs("adv_middle_Pic3Txt")
adv_middle_Pic4Txt=rs("adv_middle_Pic4Txt")
adv_middle_Pic5Txt=rs("adv_middle_Pic5Txt")
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
   adv_middle_Pic1=my_request("adv_middle_Pic1",0)
   adv_middle_Pic2=my_request("adv_middle_Pic2",0)
   adv_middle_Pic3=my_request("adv_middle_Pic3",0)
   adv_middle_Pic4=my_request("adv_middle_Pic4",0)
   adv_middle_Pic5=my_request("adv_middle_Pic5",0)
   adv_middle_Pic1url=my_request("adv_middle_Pic1url",0)
   adv_middle_Pic2url=my_request("adv_middle_Pic2url",0)
   adv_middle_Pic3url=my_request("adv_middle_Pic3url",0)
   adv_middle_Pic4url=my_request("adv_middle_Pic4url",0)
   adv_middle_Pic5url=my_request("adv_middle_Pic5url",0)
   adv_middle_Pic1Txt=my_request("adv_middle_Pic1Txt",0)
   adv_middle_Pic2Txt=my_request("adv_middle_Pic2Txt",0)
   adv_middle_Pic3Txt=my_request("adv_middle_Pic3Txt",0)
   adv_middle_Pic4Txt=my_request("adv_middle_Pic4Txt",0)
   adv_middle_Pic5Txt=my_request("adv_middle_Pic5Txt",0)
               
   if adv_middle_Pic1="" or adv_middle_Pic2="" then
       response.redirect "error.htm"
       response.end
   else
       Set rs=Server.CreateObject("ADODB.Recordset")
       sql="select * from adv_middle where adv_middle_id=1"
       rs.open sql,conn,1,3
       rs("adv_middle_Pic5")=adv_middle_Pic5
       rs("adv_middle_Pic4")=adv_middle_Pic4
       rs("adv_middle_Pic3")=adv_middle_Pic3
       rs("adv_middle_Pic1")=adv_middle_Pic1
       rs("adv_middle_Pic2")=adv_middle_Pic2
       rs("adv_middle_Pic5url")=adv_middle_Pic5url
       rs("adv_middle_Pic4url")=adv_middle_Pic4url
       rs("adv_middle_Pic3url")=adv_middle_Pic3url
       rs("adv_middle_Pic1url")=adv_middle_Pic1url
       rs("adv_middle_Pic2url")=adv_middle_Pic2url
       rs("adv_middle_Pic1Txt")=adv_middle_Pic1Txt
       rs("adv_middle_Pic2Txt")=adv_middle_Pic2Txt
       rs("adv_middle_Pic3Txt")=adv_middle_Pic3Txt
       rs("adv_middle_Pic4Txt")=adv_middle_Pic4Txt
       rs("adv_middle_Pic5Txt")=adv_middle_Pic5Txt
       rs.update
       rs.close
       set rs=nothing

      call ok("您已成功保存中间轮显广告设置！","adv_middle_set.asp")
  end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>首页轮显广告-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="adv_middle_set.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">首页轮显广告-设置　</td>
	</tr>
	<tr>
		<td colspan="2"><font color="#FF0000">图片尺寸要求： 格式:JPG&nbsp;&nbsp;&nbsp; 长:465像素&nbsp;&nbsp;&nbsp;高:160像素&nbsp;&nbsp; 
		(若不符合尺寸大小及比例可能导致不美观!)</font></td>
		</tr>
	<tr>
		<td colspan="2" class="altbg1">轮显广告图片一</td>
	</tr>
	<tr>
		<td>图片一上传：</td>
		<td>
		        <input type="text" name="adv_middle_Pic1" size="30" value="<%=adv_middle_Pic1%>"> <input type="button" value="&gt;&gt;点此上传图片" name="action1" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=adv_middle_Pic1','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')"> </td>
	</tr>
	<tr>
		<td>图片一链接网址：</td>
		<td>
		<input type="text" name="adv_middle_Pic1url" size="60" value="<%=adv_middle_Pic1url%>"></td>
	</tr>
	<tr>
		<td>图片一说明文字：</td>
		<td>
		<input type="text" name="adv_middle_Pic1Txt" size="30" value="<%=adv_middle_Pic1Txt%>" maxlength="28"></td>
	</tr>
	<tr>
		<td colspan="2" class="altbg1">轮显广告图片二</td>
	</tr>
	<tr>
		<td>图片二上传：</td>
		<td>
		        <input type="text" name="adv_middle_Pic2" size="30" value="<%=adv_middle_Pic2%>"> <input type="button" value="&gt;&gt;点此上传图片" name="action5" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=adv_middle_Pic2','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')"> </td>
	</tr>
	<tr>
		<td>图片二链接：</td>
		<td>
		<input type="text" name="adv_middle_Pic2url" size="60" value="<%=adv_middle_Pic2url%>"></td>
	</tr>
	<tr>
		<td>图片二说明文字：</td>
		<td>
		<input type="text" name="adv_middle_Pic2Txt" size="30" value="<%=adv_middle_Pic2Txt%>" maxlength="28"></td>
	</tr>
	<tr>
		<td colspan="2" class="altbg1">轮显广告图片三</td>
	</tr>
	<tr>
		<td>图片三上传：</td>
		<td>
		        <input type="text" name="adv_middle_Pic3" size="30" value="<%=adv_middle_Pic3%>"> <input type="button" value="&gt;&gt;点此上传图片" name="action6" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=adv_middle_Pic3','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')"> </td>
	</tr>
	<tr>
		<td>图片三链接：</td>
		<td>
		<input type="text" name="adv_middle_Pic3url" size="60" value="<%=adv_middle_Pic3url%>"></td>
	</tr>
	<tr>
		<td>图片三说明文字：</td>
		<td>
		<input type="text" name="adv_middle_Pic3Txt" size="30" value="<%=adv_middle_Pic3Txt%>" maxlength="28"></td>
	</tr>
	<tr>
		<td colspan="2" class="altbg1">轮显广告图片四</td>
	</tr>
<tr>
		<td>图片四上传：</td>
		<td>
		        <input type="text" name="adv_middle_Pic4" size="30" value="<%=adv_middle_Pic4%>"> 
				<input type="button" value="&gt;&gt;点此上传图片" name="action8" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=adv_middle_Pic4','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')"> </td>
	</tr>
<tr>
		<td>图片四链接：</td>
		<td>
		<input type="text" name="adv_middle_Pic4url" size="60" value="<%=adv_middle_Pic4url%>"></td>
	</tr>
<tr>
		<td>图片四说明文字：</td>
		<td>
		<input type="text" name="adv_middle_Pic4Txt" size="30" value="<%=adv_middle_Pic4Txt%>" maxlength="28"></td>
	</tr>
	<tr>
		<td colspan="2" class="altbg1">轮显广告图片五</td>
	</tr>
<tr>
		<td>图片五上传：</td>
		<td>
		        <input type="text" name="adv_middle_Pic5" size="30" value="<%=adv_middle_Pic5%>"> 
				<input type="button" value="&gt;&gt;点此上传图片" name="action7" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=adv_middle_Pic5','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=yes,resizable=yes,width=400,height=100')"> </td>
	</tr>
<tr>
		<td>图片五链接：</td>
		<td>
		<input type="text" name="adv_middle_Pic5url" size="60" value="<%=adv_middle_Pic5url%>"></td>
	</tr>
<tr>
		<td>图片五说明文字：</td>
		<td>
		<input type="text" name="adv_middle_Pic5Txt" size="30" value="<%=adv_middle_Pic5Txt%>" maxlength="28"></td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		   <input type="submit" value=" 保存设置 " name="Submit1"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

