<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
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
    root_info_aboutus =my_request("content",0)
     
    ErrMsg=""
    if root_info_aboutus="" then
    	FoundErr=True
	    ErrMsg=ErrMsg & "<li>[关于我们] 的内容不能为空！</li>"
    end if
    if FoundErr<>True then
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from root_info where id=1"
        rs.open sql,conn,1,3
        rs("root_info_aboutus") =root_info_aboutus
        rs.update
        rs.close
        set rs=nothing
        call ok("您已成功保存[关于我们]设置！","root_aboutus_set.asp")
    else
        call WriteErrMsg(ErrMsg)
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>关于我们-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="Root_AboutUs_Set.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td class="header" colspan="2">关于我们- 设置</td>
	</tr>
	<tr>
		<td>关于我们：</td>
		<td>
		    <!--#include file="editor/editor.asp"-->
            <script language="javascript">
                document.write ('<iframe src="Root_AboutUs_TxtBox.asp" id="message" width="95%" height="400"></iframe>')
                frames.message.document.designMode = "On";
            </script>
</td>
	</tr>
	
	<tr>
		<td>　</td>
		<td><input type="submit" value="  提  交  " name="B1" onclick="document.form1.Content.value = frames.message.document.body.innerHTML;">&nbsp; 
		    <input type="reset" value="重置" name="B2"><input type="hidden" name="Content" value>
        </td>
	</tr>
	</form>
 </tbody>
</table>
<br>
</body>

</html>

