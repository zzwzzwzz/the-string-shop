<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<title>管理员登录</title>
<link rel="stylesheet" type="text/css" href="style.css">
<SCRIPT language=JavaScript>
<!--
function check(form)
{

 if (form.login_name.value=="")
 {
 alert('用户名不能为空!');
 document.MM_returnValue=false;
 }
else if(form.login_pass.value=="")
 {
   alert('密码不能为空!');
   document.MM_returnValue=false;
  }
else if(form.codeid.value=="")
  {
   alert('验证码不能为空!');
   document.MM_returnValue=false;
  }

else {
   document.MM_returnValue=true;
   }
}

//-->
</SCRIPT>
</head>

<body>
<br><br><br><br><br>
<table cellspacing="1" cellpadding="4" width="30%" class="tableborder" align="center">
<tbody class="altbg2">
<FORM name=manage action=Admin_LoginCheck.asp method=post>
	<tr>
		<td colspan="2" class="header">管理员-登陆</td>
	</tr>
	<tr>
		<td>用户名：</td>
		<td><INPUT size=20 name=login_name></td>
	</tr>
	<tr>
		<td>密&nbsp;&nbsp;&nbsp; 码：</td>
		<td><INPUT type=password size=20 name=login_pass></td>
	</tr>
	<tr>
		<td>验证码：</td>
		<td><INPUT size=8 name=codeid><img src="../include/admincheckcode.asp"></td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		    <INPUT onclick="check(this.form);return document.MM_returnValue" type=submit value=提交 name=B1>&nbsp;&nbsp;
			<INPUT type=reset value=重置 name=B2>
	    </td>
	</tr>
</tbody>
</table>

</body>

</html>

 
