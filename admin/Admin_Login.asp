<html>

<head>
<link rel="shortcut icon" href="/IMAGES/favicon.ico">
<meta http-equiv="Content-Language" content="zh-cn">
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<title>����Ա��¼</title>
<link rel="stylesheet" type="text/css" href="style.css">
<SCRIPT language=JavaScript>
function check(form)
{

 if (form.login_name.value=="")
 {
 alert('�û�������Ϊ��!');
 document.MM_returnValue=false;
 }
else if(form.login_pass.value=="")
 {
   alert('���벻��Ϊ��!');
   document.MM_returnValue=false;
  }

else {
   document.MM_returnValue=true;
   }
}

</SCRIPT>
</head>

<body>
<br><br><br><br><br><br><br><br><br><br><br><br>
<table cellspacing="1" cellpadding="5" width="20%" class="tableborder" align="center">
<tbody class="altbg2">
<FORM name=manage action=Admin_LoginCheck.asp method=post>
	<tr>
		<td colspan="2" class="header">����Ա-��¼</td>
	</tr>
	<tr>
		<td align="right">�û�����</td>
		<td><INPUT size=20 name=login_name></td>
	</tr>
	<tr>
		<td align="right">��&nbsp;&nbsp;&nbsp;�룺</td>
		<td><INPUT type=password size=20 name=login_pass></td>
	</tr>
	<tr>
		<td>��</td>
		<td>
		    <INPUT onclick="check(this.form);return document.MM_returnValue" type=submit value=�ύ name=B1>&nbsp;&nbsp;
			<INPUT type=reset value=���� name=B2>
	    </td>
	</tr>
</tbody>
</table>

</body>

</html>