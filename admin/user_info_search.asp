<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=3
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ա-��Ա��Ϣ-����</title>
<link rel="stylesheet" type="text/css" href="style.css">

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="user_info_list.asp" method="get">
<input type="hidden" name="txt_type_hidden" value="1">
	<tr>
		<td colspan="2" class="header">��Ա��Ϣ-�߼�����</td>
	</tr>
	<tr>
		<td>�û�����</td>
		<td><input type="text" name="KeyWord" size="40"></td>
	</tr>
	<tr>
		<td>��ʵ������</td>
		<td><input type="text" name="prod_info_no" size="40"></td>
	</tr>
	<tr>
		<td>�������䣺</td>
		<td><input type="text" name="prod_info_detail" size="40"></td>
	</tr>
	<tr>
		<td>��ϵ�绰��</td>
		<td><input type="text" name="prod_info_detail4" size="40"></td>
	</tr>
	<tr>
		<td>�������룺</td>
		<td><input type="text" name="prod_info_detail3" size="40"></td>
	</tr>
		<tr>
		<td>ע��ʱ�䣺</td>
		<td><input type="radio" value="1" name="spec">���� 
		<input type="radio" value="0" name="spec">���� 
		<input type="radio" value="2" name="spec">һ���� 
		<input type="radio" value="21" name="spec">һ���� 
		<input type="radio" value="22" name="spec">ȫ��</td>
	</tr>
	<tr>
		<td>����½ʱ�䣺</td>
		<td><input type="radio" value="11" name="spec">���� 
		<input type="radio" value="01" name="spec">���� 
		<input type="radio" value="23" name="spec">һ���� 
		<input type="radio" value="24" name="spec">һ���� 
		<input type="radio" value="25" name="spec" checked>ȫ��</td>
	</tr>
	<tr>
		<td>��½������</td>
		<td><input type="text" name="prod_info_UserPriceMin1" size="6">���� 
		<input type="text" name="prod_info_UserPriceMin0" size="6">��</td>
	</tr>
	<tr>
		<td>��Ա״̬��</td>
		<td><input type="radio" value="11" name="new">���� 
		<input type="radio" value="01" name="new">����/�����</td>
	</tr>
	<tr>
		<td>��</td>
		<td><input type="submit" value="  �� ��  " name="B1"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

 
