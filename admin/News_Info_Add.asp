<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=4
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
    news_info_title = my_request("news_info_title",0)
    news_info_type  = my_request("news_info_type",1)   '0=网址链接 1=内容填充
    if news_info_type=0 then news_info_content = my_request("news_info_content",0)
    if news_info_type=1 then news_info_content = my_request("Content",0)
    ErrMsg=""
    if news_info_title="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>公告信息标题不能为空！</li>"
    end if
    if news_info_content="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>公告信息内容不能为空！</li>"
    end if
    if FoundErr<>True then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from news_info"
        rs.open sql,conn,1,3
        rs.addnew
        rs("news_info_title")   = news_info_title
        rs("news_info_type")    = news_info_type
        rs("news_info_content") = news_info_content
        rs("news_info_addtime") = now()
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功添加了一条网上公告信息！","news_info_List.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文章信息-添加</title>
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
<form action="news_info_add.asp" method="post" name="form1" onsubmit="return checkdata();">
<input type="hidden" name="action" value="save">
	<tr>
		<td colspan="2" class="header">文章信息-添加</td>
	</tr>
	<tr>
		<td>标题：</td>
		<td><input type="text" name="news_info_title" size="40"></td>
	</tr>
	<tr>
		<td>内容类型：</td>
		<td>
		<input type="radio" value="0" name="news_info_type" checked onClick='showlist("a");'>网址链接
			<input type="radio" value="1" name="news_info_type" onClick='showlist("b");'>内容填充</td>
	</tr>
	<tr id="linkimg2" style='display:""'>
		<td>链接网址：</td>
		<td><input type="text" name="news_info_content" size="40"></td>
	</tr>
	<tr id="linkimg" style='display:none'>
		<td height="23">公告内容：</td>
		<td height="23">
		    <!--#include file="editor/editor.asp"-->
            <script language="javascript">
                document.write ('<iframe src="News_TxtBox.asp" id="message" width="95%" height="200"></iframe>')
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

