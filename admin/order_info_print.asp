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
<title>订单-订单信息-打印</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="order_info_PrintDetail.asp" method="post">
	<tr>
		<td colspan="2" class="header">订单打印</td>
	</tr>
	<tr>
		<td>请输入订单号：</td>
		<td><input type="text" name="order_info_no" size="30"></td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="生成打印件" name="B1"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 
