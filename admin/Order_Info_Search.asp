<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=2
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����-������Ϣ-����</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="order_info_list.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">��������</td>
	</tr>
	<tr>
		<td>���ݶ���״̬��</td>
		<td><select name="search_order_CheckStates">
		<option value=''>��ѡ�񶩵�״̬</option>
		<option value="0">�¶���</option>
		<option value="1">��Ա����ȡ��</option>
		<option value="2">��Ч������ȡ��</option>
		<option value="3">��ȷ�ϣ�������</option>
		<option value="4">�ѷ��������ջ�</option>
		<option value="5">����֧���ɹ�</option>
		<option value="6">�������</option>
		</select></td>
	</tr>
	<tr>
		<td>���ݶ����ţ�</td>
		<td><input type="text" name="search_order_no" size="30"></td>
	</tr>
	<tr>
		<td>���ݶ�����ԱID��</td>
		<td><input type="text" name="search_order_UserName" size="30"></td>
	</tr>
	<tr>
		<td>�����ջ���������</td>
		<td><input type="text" name="search_order_RealName" size="30"></td>
	</tr>
	<tr>
		<td>���ݵ������䣺</td>
		<td><input type="text" name="search_order_email" size="30"></td>
	</tr>
	<tr>
		<td>������ϵ�绰��</td>
		<td><input type="text" name="search_order_mobile" size="30"></td>
	</tr>
	<tr>
		<td>������ϵ��ַ��</td>
		<td><input type="text" name="search_order_address" size="30"></td>
	</tr>
	<tr>
		<td>�����������룺</td>
		<td><input type="text" name="search_order_zip" size="30"></td>
	</tr>
	<tr>
		<td>���ݶ���ʱ�䣺</td>
		<td><input type="radio" value="1" name="search_order_BuyTime">����&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="2" name="search_order_BuyTime">����&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="7" name="search_order_BuyTime">һ����&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="30" name="search_order_BuyTime">һ����&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="" checked name="search_order_BuyTime">ȫ��&nbsp;&nbsp; </td>
	</tr>
	<tr>
		<td>��</td>
		<td>
		   <input type="submit" value="��ʼ����" name="Submit1"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 
