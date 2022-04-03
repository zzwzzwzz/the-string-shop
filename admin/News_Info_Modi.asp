<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=4
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""News_Info_List.asp"";</script>")
  response.end
end if

Set rs= Server.CreateObject("ADODB.Recordset")
sql="select news_info_title,news_info_type,news_info_content from news_info where id="&id
rs.open sql,conn,1,1
news_info_title	  = rs(0)
news_info_type	  = rs(1)
if news_info_type=0 then 
    news_info_content = rs(2)
else
    news_info_content=""
end if
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id				= my_request("id",1)
    news_info_title = my_request("news_info_title",0)
    news_info_type  = my_request("news_info_type",1)   '0=网址链接 1=内容填充
    if news_info_type=0 then news_info_content = my_request("news_info_content",0)
    if news_info_type=1 then news_info_content = my_request("Content",0)
    ErrMsg=""
    if id="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>文章文章信息ID不能为空！</li>"
    end if
    if news_info_title="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>文章信息标题不能为空！</li>"
    end if
    if news_info_content="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>文章文章信息内容不能为空！</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from news_info where id="&id
        rs.open sql,conn,1,3
        rs("news_info_title")   = news_info_title
        rs("news_info_type")    = news_info_type
        rs("news_info_content") = news_info_content
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功编辑了一条文章信息！","news_info_List.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文章信息-编辑</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript">
<!--
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

function checkdata()
{
if (document.form1.viewhtml.checked == true)
	{
	  alert("对不起，请取消“查看HTML源代码”后再添加！")
	  document.form1.viewhtml.focus()
	  return false
	 }
}
//-->
</script>
<script src="Editor/edit.js" type="text/javascript"></script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form action="News_Info_Modi.asp" method="post" name="form1" onsubmit="return checkdata();">
<input type="hidden" name="action" value="save">
<input type="hidden" name="id" value="<%=id%>"> 
	<tr>
		<td colspan="2" class="header">文章信息-编辑</td>
	</tr>
	<tr>
		<td>标题：</td>
		<td>
		<input type="text" name="news_info_title" size="40" value="<%=news_info_title%>"></td>
	</tr>
	<tr>
		<td>内容类型：</td>
		<td>
		<input type="radio" value="0" name="news_info_type" checked onClick='showlist("a");' <%if news_info_type=0 then response.write "checked" %>>网址链接
			<input type="radio" value="1" name="news_info_type" onClick='showlist("b");' <%if news_info_type=1 then response.write "checked" %>>内容填充</td>
	</tr>
	<tr id="linkimg2" <%if news_info_type=0 then%>style='display:""'<%else%>style='display:none'<%end if%>>
		<td>链接网址：</td>
		<td>
		<input type="text" name="news_info_content" size="40" value="<%=news_info_content%>"></td>
	</tr>
	<tr id="linkimg" <%if news_info_type=1 then%>style='display:""'<%else%>style='display:none'<%end if%>>
		<td height="23"文章内容：</td>
		<td height="23">
		    <!--#include file="editor/editor.asp"-->
            <script language="javascript">
                document.write ('<iframe src="News_TxtBox.asp?id=<%=id%>&action=modify" id="message" width="95%" height="200"></iframe>')
                frames.message.document.designMode = "On";
            </script>
        </td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		    <input type="submit" value="  提  交  " name="B1" onclick="document.form1.Content.value = frames.message.document.body.innerHTML;">&nbsp; 
		    <input type="reset" value="重置" name="B2"><input type="hidden" name="Content" value>
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

