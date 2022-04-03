<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=3
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
email =my_request("email",0)

user_info_id=my_request("user_info_id",0)
if user_info_id<>"" then
    pp=ubound(split(user_info_id,","))+1 '判断数组id中共有几维
    for v=1 to pp
        id=request("user_info_id")(v)     
        set rs=conn.execute ("select user_info_email from [user_info] where user_info_id="&id)
        user_info_email=rs(0)
        rs.close
        set rs=nothing
        email=user_info_email&","&email
    next
    intHowLong=len(email)
    email=left(email,intHowLong-1)

end if

action=my_request("action",0)
if action="save" then
    call save()
end if

Sub SendAction(subject, mailaddress, email, sender, content, fromer,username,password) 
    Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象 
    jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值j 
    jmail.logging = true '启用邮件日志 
    jmail.Charset = "GB2312" '邮件的文字编码为国标 
    jmail.ContentType = "text/html" '邮件的格式为HTML格式 
    jmail.AddRecipient email '邮件收件人的地址 
    jmail.From = fromer '发件人的E-MAIL地址 
    jmail.MailServerUserName = username '登录邮件服务器所需的用户名 
    jmail.MailServerPassword = password '登录邮件服务器所需的密码 
    jmail.Subject = subject '邮件的标题 
    jmail.Body = content '邮件的内容 
    jmail.Priority = 1 '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值 
    jmail.Send(mailaddress) '执行邮件发送（通过邮件服务器地址） 
    jmail.Close() '关闭对象 
End Sub

sub save()    
    sql="select root_email_server,root_email_UserName,root_email_PassWord from root_email"
    set rs=conn.execute (sql)
    root_email_server   =rs(0)
    root_email_UserName =rs(1)
    root_email_PassWord =rs(2)
    rs.close
    set rs=nothing

    email_receive_man=my_request("email_receive_man",0)  
    email_info_title =my_request("email_info_title",0)
    email_info_detail=my_request("Content",0)
    if email_receive_man="" or email_info_title="" or email_info_detail="" then
        call error()
    else
        a=split(email_receive_man,",")
        for i=0 to ubound(a)
            call SendAction(email_info_title, root_email_server, a(i), root_email_UserName, email_info_detail, root_email_UserName,root_email_UserName,root_email_PassWord)
            response.write "<li>"&a(i)&"： 发送成功!</li>" 
        next
        call ok("您已成功发送了邮件信息！","user_email_send.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>会员-邮件-群发</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script src="Editor/edit.js" type="text/javascript"></script>
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="user_email_send.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">向会员发送邮件</td>
	</tr>
	<tr>
		<td colspan="2"><font color="#ff3300">特别说明：</font><font color="#999999">请确认你网站所在的服务器支持<b>Jmail</b>邮件发送组件，否则以下各项功能均无法正常使用！若不清楚是否支持<b>Jmail</b>，请向服务商咨询！</font></td>
	</tr>
	<tr>
		<td>收件人(Email地址)：<font color="#999999"><br>
		多个收件人信箱请用逗号&quot;，&quot;隔开</font></td>
		<td><textarea rows="6" name="email_receive_man" cols="60"><%=email%></textarea></td>
	</tr>
	<tr>
		<td>邮件主题：</td>
		<td><input type="text" name="email_info_title" size="62"></td>
	</tr>
	<tr>
		<td>邮件内容：</td>
		<td> <!--//商品介绍//-->
		   <!--#include file="editor/editor.asp"-->
           <script language="javascript">
           document.write ('<iframe src="email_txtbox.asp" id="message" width="90%" height="300"></iframe>')
           frames.message.document.designMode = "On";
           </script>
        </td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="开始群发" name="Submit1" onclick="document.form1.Content.value = frames.message.document.body.innerHTML;">&nbsp; 
		   <input type="reset" value="重置" name="B2">
		   <input type="hidden" name="Content" value>		</td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 
