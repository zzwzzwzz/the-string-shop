<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=8
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""root_model_List.asp"";</script>")
  response.end
end if

sql="select root_model_name,root_model_css,id from root_model where id="&id
set rs=conn.execute (sql)
root_model_name=rs(0)
root_model_css=rs(1)
id=rs(2)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
   call save()
end if

sub save()
    id=my_request("id",1)
    root_model_name=my_request("root_model_name",0)
    root_model_css=my_request("root_model_css",0)
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
        sql="select * from root_model where id="&id
        rs.open sql,conn,1,3
        rs("root_model_name")=root_model_name
        rs("root_model_css")=root_model_css
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功更新了一条网站模板信息！","root_model_list.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>模板-更新</title>
<link rel="stylesheet" type="text/css" href="style.css">
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
<form name="form1" action="Root_Model_Modi.asp" method="post">
<input type="hidden" name="action" value="save"> 
<input type="hidden" name="id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">模板-更新</td>
	</tr>
	<tr>
		<td>模板名称：</td>
		<td>
		    <input type="text" name="root_model_name" size="20" value="<%=root_model_name%>"></td>
	</tr>
	<tr>
		<td>模板样式表-文件名：</td>
		<td>
		    <input type="text" name="root_model_css" size="20" value="<%=root_model_css%>">.css<font color="#808080">&nbsp;&nbsp;
			<br>
			请确认你已将此文件放到了style目录下了;<br>
			该模板用到的图片文件包也请一并放到style目录下;</font></td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		   <input type="submit" value="  提  交  " name="Submit1">&nbsp; 
		   <input type="reset" value="  重  置  " name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

