<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=8
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
action=my_request("action",0)
if action="save" then
   call save()
end if

sub save()
    root_model_name=my_request("root_model_name",0)
    root_model_css=my_request("root_model_css",0)
    root_model_ispic=my_request("root_model_ispic",0)
    if root_model_ispic=1 then root_model_pic=my_request("root_model_pic",0)

    ErrMsg=""
    if root_model_name="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>模板名称不能为空！</li>"
    end if
    if root_model_css="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>样式表文件名不能为空！</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from root_model"
        rs.open sql,conn,1,3
        rs.addnew
        rs("root_model_name")=root_model_name
        rs("root_model_css")=root_model_css
        if root_model_ispic=1 then
        	rs("root_model_pic")=root_model_pic
        end if
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功添加了一条网站模板信息！","root_model_list.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>模板-添加</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
<script language = "JavaScript">  
function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="none";
  }
  else
  {
   linkimg.style.display="";
  }
}
</script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="Root_Model_Add.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">模板-添加</td>
	</tr>
	<tr>
		<td colspan="2"><font color="#808080">[说明]:在添加模板前,请您</font><font color="#FF6600">先将样式表文件及该模板用到的图片文件包放到style目录下,</font><font color="#808080">否则无效!</font></td>
	</tr>
	<tr>
		<td>模板名称：</td>
		<td>
		    <input type="text" name="root_model_name" size="20"></td>
	</tr>
	<tr>
		<td>模板样式表-文件名：</td>
		<td>
		    <input type="text" name="root_model_css" size="20">.css<font color="#808080">&nbsp;&nbsp;
			<br>
			请确认你已将此文件放到了style目录下了;<br>
			该模板用到的图片文件包也请一并放到style目录下;</font></td>
	</tr>
	<tr>
		<td>是否有模板首页截图：</td>
		<td> 
		<input type="radio" value="1" name="root_model_ispic" onClick='showlist("b");' checked>是&nbsp;&nbsp;&nbsp;
		<input type="radio" value="0" name="root_model_ispic" onClick='showlist("a");'>否&nbsp;
		</td>
	</tr>
	<tr id="linkimg">
		<td>模板首页截图：</td>
		<td><input type="text" name="root_model_pic" size="40">
		        <input type="button" value="&gt;&gt;点此上传图" name="action" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=root_model_pic','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=400,height=100')">
		</td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		   <input type="submit" value="  提  交  " name="Submit1">&nbsp; 
		   <input type="reset" value="重置" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

