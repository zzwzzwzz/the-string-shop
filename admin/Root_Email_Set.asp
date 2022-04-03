<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=0
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_email_server,root_email_email,root_email_PassWord from root_email where root_email_id=1"
rs.open sql,conn,1,1
root_email_server       =rs(0)
root_email_email        =rs(1)
root_email_PassWord     =rs(2)
rs.close
set rs=nothing

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    root_email_server=my_request("root_email_server",0)
    root_email_email=my_request("root_email_email",0)
    root_email_PassWord=my_request("root_email_PassWord",0)
            
    if root_email_server="" or root_email_email="" or root_email_PassWord="" then
        response.redirect "error.htm"
        response.end
    else
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select * from root_email where root_email_id=1"
        rs.open sql,conn,1,3
        rs("root_email_server")=root_email_server
        rs("root_email_email")=root_email_email  
        rs("root_email_PassWord")=root_email_PassWord 
        rs.update
        rs.close
        set rs=nothing


        call ok("您已成功保存发送邮件设置！","root_email_set.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>基本-邮件-设置</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="root_email_set.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">发送邮件设置</td>
	</tr>
	<tr>
		<td colspan="2"><font color="#FF3300">特别说明：</font><font color="#999999">请确认你网站所在的服务器支持<b>Jmail</b>邮件发送组件，否则以下各项功能均无法正常使用！若不清楚是否支持<b>Jmail</b>，请向服务商咨询！</font></td>
	</tr>
	<tr>
		<td><b>发信邮件服务器：</b><br>
		用于发信的SMTP服务器</td>
		<td>
		<input type="text" name="root_email_server" size="30" value="<%=root_email_server%>"></td>
	</tr>
	<tr>
		<td><b>发信邮箱：</b><br>
		用于发信的邮箱名称，建议不要使用重要邮箱作为发信邮箱</td>
		<td>
		<input type="text" name="root_email_email" size="30" value="<%=root_email_email%>"></td>
	</tr>
	<tr>
		<td><b>邮箱登录密码：</b><br>
		填写发信邮箱的密码，特别提示：此密码以未加密方式保存在数据库中，因此，建议您设置一个专用的发信邮箱，不要使用重要的邮箱作为发信邮箱</td>
		<td>
		<input type="password" name="root_email_PassWord" size="30" value="<%=root_email_PassWord%>"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="保存设置" name="Submit1">&nbsp;&nbsp; 
		   <input type="reset" value="重置" name="B2">
		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 
