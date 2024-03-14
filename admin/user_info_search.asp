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
<title>会员-会员信息-搜索</title>
<link rel="stylesheet" type="text/css" href="style.css">

</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="user_info_list.asp" method="get">
<input type="hidden" name="txt_type_hidden" value="1">
	<tr>
		<td colspan="2" class="header">会员信息-高级搜索</td>
	</tr>
	<tr>
		<td>用户名：</td>
		<td><input type="text" name="KeyWord" size="40"></td>
	</tr>
	<tr>
		<td>真实姓名：</td>
		<td><input type="text" name="prod_info_no" size="40"></td>
	</tr>
	<tr>
		<td>电子邮箱：</td>
		<td><input type="text" name="prod_info_detail" size="40"></td>
	</tr>
	<tr>
		<td>联系电话：</td>
		<td><input type="text" name="prod_info_detail4" size="40"></td>
	</tr>
	<tr>
		<td>邮政编码：</td>
		<td><input type="text" name="prod_info_detail3" size="40"></td>
	</tr>
		<tr>
		<td>注册时间：</td>
		<td><input type="radio" value="1" name="spec">今天 
		<input type="radio" value="0" name="spec">昨天 
		<input type="radio" value="2" name="spec">一周内 
		<input type="radio" value="21" name="spec">一月内 
		<input type="radio" value="22" name="spec">全部</td>
	</tr>
	<tr>
		<td>最后登陆时间：</td>
		<td><input type="radio" value="11" name="spec">今天 
		<input type="radio" value="01" name="spec">昨天 
		<input type="radio" value="23" name="spec">一周内 
		<input type="radio" value="24" name="spec">一月内 
		<input type="radio" value="25" name="spec" checked>全部</td>
	</tr>
	<tr>
		<td>登陆次数：</td>
		<td><input type="text" name="prod_info_UserPriceMin1" size="6">次至 
		<input type="text" name="prod_info_UserPriceMin0" size="6">次</td>
	</tr>
	<tr>
		<td>会员状态：</td>
		<td><input type="radio" value="11" name="new">正常 
		<input type="radio" value="01" name="new">锁定/审核中</td>
	</tr>
	<tr>
		<td>　</td>
		<td><input type="submit" value="  搜 索  " name="B1"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

 
