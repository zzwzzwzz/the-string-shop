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
  response.write("<script>alert(""参数错误!"");location.href=""link_info_List.asp"";</script>")
  response.end
end if

sql="select * from link_info where id="&id
set rs=conn.execute (sql)
link_info_type=rs("link_info_type")
link_info_url=rs("link_info_url")
if link_info_type=0 then
    link_info_title=rs("link_info_detail")
end if
if link_info_type=1 then
    link_info_LogoPic=rs("link_info_detail")
end if
link_info_indexshow=rs("link_info_indexshow")
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then call save()

sub save()
    id=my_request("id",1)
    link_info_type=my_request("link_info_type",1)
    if link_info_type=0 then link_info_detail=my_request("link_info_title",0)
    if link_info_type=1 then link_info_detail=my_request("link_info_LogoPic",0)
    link_info_url=my_request("link_info_url",0)
    link_info_IndexShow=my_request("link_info_IndexShow",1)
    ErrMsg=""
    if link_info_type="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>类型不能为空！</li>"
    end if
    if link_info_url="" or link_info_url="http://" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>链接网址不能为空！</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from link_info where id="&id
        rs.open sql,conn,1,3
        rs("link_info_type")=link_info_type
        rs("link_info_url")=link_info_url
        rs("link_info_detail")=link_info_detail
        rs("link_info_IndexShow")=link_info_IndexShow
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功修改了一条友情链接信息！","link_info_list.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>友情链接-友情链接信息-修改</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
<script language = "JavaScript">  
function showlist(dd)
{
  if(dd=="a")
  {
   linkimg.style.display="none";
   linkimg2.style.display="";
  }
  else
  {
   linkimg.style.display="";
   linkimg2.style.display="none";
  }
}
</script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="link_info_modi.asp" method="post">
<input type="hidden" name="action" value="save"> 
<input type="hidden" name="id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">友情链接-修改</td>
	</tr>
	<tr>
		<td>链接类型：</td>
		<td>
		<input type="radio" value="0" name="link_info_type" checked onClick='showlist("a");' <%if link_info_type=0 then response.write "checked" %>>文字链接&nbsp; 
		<input type="radio" value="1" name="link_info_type" onClick='showlist("b");' <%if link_info_type=1 then response.write "checked" %>>图标链接</td>
	</tr>
	<tr id="linkimg2" <%if link_info_type=0 then%>style='display:""'<%else%>style='display:none'<%end if%>>
		<td>链接文字：</td>
		<td>
		    <input type="text" name="link_info_title" size="40" value="<%=link_info_title%>"></td>
	</tr>
	<tr id="linkimg" <%if link_info_type=1 then%>style='display:""'<%else%>style='display:none'<%end if%>>
		<td>链接图标：</td>
		<td><%if link_info_type=1 and link_info_logoPic<>"" then%>原图预览：<img src="../uploadpic/<%=link_info_logoPic%>"><br><%end if%>
		<input type="text" name="link_info_LogoPic" size="40" value="<%=link_info_LogoPic%>">
		        <input type="button" value="&gt;&gt;点此上传图标" name="action" onclick="javascript:openWin('Njj_Pic_Upload.asp?Fname=link_info_LogoPic','upload','toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=yes,width=400,height=100')">
		</td>
	</tr>
	<tr>
		<td>链接网址：</td>
		<td>
		    <input type="text" name="link_info_url" size="40" value="<%=link_info_url%>"></td>
	</tr>
	<tr>
		<td>是否-首页显示：</td>
		<td>
		<input type="radio" value="0" name="link_info_IndexShow" <%if link_info_IndexShow=0 then response.write "checked" %>>是&nbsp;&nbsp;&nbsp;
		<input type="radio" value="1" name="link_info_IndexShow" <%if link_info_IndexShow=1 then response.write "checked" %>>否</td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		   <input type="submit" value="提交" name="Submit1">&nbsp; 
		   <input type="reset" value="重置" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

