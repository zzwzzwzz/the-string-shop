<!--#include file="admin_check.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ർ��</title>
<link rel="stylesheet" type="text/css" href="style.css">
<base target="main">
</head>

<body>

<table border="0" width="100%" cellpadding="4" style="border:1px solid #cccccc; border-collapse: collapse; padding-left:4px; padding-right:4px; padding-top:1px; padding-bottom:1px" bgcolor="#FFFFFF">

	<tr>
		<td align="center"><a href="Right.asp" target="main">��̨��ҳ</a> <font color="#999999">| </font>&nbsp;<a href="Admin_LoginOut.asp" target=_parent>�˳�ϵͳ</a></td>
	</tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td height="1"></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">��������</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="Root_Info_Set.asp">��������</a></li>
		   <li><a target="main" href="Root_Model_list.asp">��վģ��</a></li>
		   <li><a target="main" href="Root_Option_Set.asp">����ѡ��</a></li>
		   <li><a href="Root_NetPay_Set.asp" target="main">֧����ʽ</a></li>
		   <li><a href="Root_Deliver_Set.asp" target="main">�ͻ���ʽ</a><br></li>
		   <li><a href="Root_Vote_set.asp" target="main">ͶƱ����</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">��Ʒ����</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="Prod_Class_List.asp">������</a></li>
		   <li><a target="main" href="Product_Info_List.asp">��Ʒ����</a> | <a target="main" href="Product_Info_Add.asp">����</a></li>
		   <li><a target="main" href="Product_Info_Search.asp">������Ʒ</a></li>
		   <li><a target="main" href="Product_kucun_list.asp">������</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">��������</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a href="Order_info_List.asp" target="main">ȫ������</a> | <a target="main" href="Order_info_search.asp">����</a></li>
		   <li><a href="Order_info_recycle.asp" target="main">������ԭ</a></li>
		   <li><a href="Order_info_SaleCount.asp" target="main">����ͳ��</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">��Ա����</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a href="user_info_list.asp" target="main">��Ա��Ϣ����</a></li>
		   <li><a href="user_info_search.asp" target="main">��Ա�߼�����</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">���¹���</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a href="news_info_list.asp" target="main">���¹���</a> | <a target="main" href="News_Info_Add.asp">����</a></li>
		   <li><a href="help_info_list.asp" target="main">��������</a> | <a target="main" href="help_info_add.asp">����</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">���Թ���</td>
	</tr>
	<tr>
		<td class="altbg2">
		    <li><a target="main" href="GB_Info_List.asp">�������Թ���</a></li>
  			<li><a target="main" href="Prod_Review_List.asp"> ��Ʒ���۹���</a></li>		   
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">Ȩ�޹���</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a href="admin_info_add.asp" target="main">������Ա����</a></li>
		   <li><a href="admin_info_list.asp" target="main">������Ա����</a></li>
		   <li><a href="admin_info_PassWordModiByUserName.asp?admin_info_UserName=<%=session("admin_info_UserName")%>" target="main">���������޸�</a></li>
		</td>
	</tr>
	</tbody>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td></td>
	</tr>
</table>

<table border="1" width="100%" cellpadding="4" style="border-collapse: collapse; padding-left:4px; padding-right:4px; padding-top:1px; padding-bottom:1px" bgcolor="#ffffff" bordercolor="#cccccc">
	<tr>
		<td align="center"><font color="#999999">�����������ϵͳ<br>
		����ߣ�30818103</font></a></td>
	</tr>
</table>
</body>
</html>