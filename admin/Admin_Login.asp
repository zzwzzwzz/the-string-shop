
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

 
