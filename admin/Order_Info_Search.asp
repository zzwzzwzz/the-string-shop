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
<title>订单-订单信息-搜索</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="order_info_list.asp" method="post">
<input type="hidden" name="action" value="save"> 
	<tr>
		<td colspan="2" class="header">订单搜索</td>
	</tr>
	<tr>
		<td>根据订单状态：</td>
		<td><select name="search_order_CheckStates">
		<option value=''>请选择订单状态</option>
		<option value="0">新订单</option>
		<option value="1">会员自行取消</option>
		<option value="2">无效单，已取消</option>
		<option value="3">已确认，待付款</option>
		<option value="4">已发货，待收货</option>
		<option value="5">在线支付成功</option>
		<option value="6">订单完成</option>
		</select></td>
	</tr>
	<tr>
		<td>根据订单号：</td>
		<td><input type="text" name="search_order_no" size="30"></td>
	</tr>
	<tr>
		<td>根据订单会员ID：</td>
		<td><input type="text" name="search_order_UserName" size="30"></td>
	</tr>
	<tr>
		<td>根据收货人姓名：</td>
		<td><input type="text" name="search_order_RealName" size="30"></td>
	</tr>
	<tr>
		<td>根据电子邮箱：</td>
		<td><input type="text" name="search_order_email" size="30"></td>
	</tr>
	<tr>
		<td>根据联系电话：</td>
		<td><input type="text" name="search_order_mobile" size="30"></td>
	</tr>
	<tr>
		<td>根据联系地址：</td>
		<td><input type="text" name="search_order_address" size="30"></td>
	</tr>
	<tr>
		<td>根据邮政编码：</td>
		<td><input type="text" name="search_order_zip" size="30"></td>
	</tr>
	<tr>
		<td>根据订购时间：</td>
		<td><input type="radio" value="1" name="search_order_BuyTime">今天&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="2" name="search_order_BuyTime">昨天&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="7" name="search_order_BuyTime">一周内&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="30" name="search_order_BuyTime">一月内&nbsp;&nbsp;&nbsp; 
		<input type="radio" value="" checked name="search_order_BuyTime">全部&nbsp;&nbsp; </td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		   <input type="submit" value="开始搜索" name="Submit1"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>
 
