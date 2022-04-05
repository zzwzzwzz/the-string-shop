<!--#include file="admin_check.asp"-->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>左侧导航</title>
<link rel="stylesheet" type="text/css" href="style.css">
<base target="main">
</head>

<body>

<table border="0" width="100%" cellpadding="4" style="border:1px solid #cccccc; border-collapse: collapse; padding-left:4px; padding-right:4px; padding-top:1px; padding-bottom:1px" bgcolor="#FFFFFF">

	<tr>
		<td align="center"><a href="Right.asp" target="main">后台首页</a> <font color="#999999">| </font>&nbsp;<a href="Admin_LoginOut.asp" target=_parent>退出系统</a></td>
	</tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" height="8">
	<tr>
		<td height="1"></td>
	</tr>
</table>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
	<tr>
		<td class="header">基本设置</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="Root_Info_Set.asp">基本资料</a></li>
		   <li><a target="main" href="Root_Model_list.asp">网站模板</a></li>
		   <li><a target="main" href="Root_Option_Set.asp">参数选项</a></li>
		   <li><a href="Root_NetPay_Set.asp" target="main">支付方式</a></li>
		   <li><a href="Root_Deliver_Set.asp" target="main">送货方式</a><br></li>
		   <li><a href="Root_Vote_set.asp" target="main">投票设置</a></li>
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
		<td class="header">商品管理</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="Prod_Class_List.asp">类别管理</a></li>
		   <li><a target="main" href="Product_Info_List.asp">商品管理</a> | <a target="main" href="Product_Info_Add.asp">添加</a></li>
		   <li><a target="main" href="Product_Info_Search.asp">搜索商品</a></li>
		   <li><a target="main" href="Product_kucun_list.asp">库存管理</a></li>
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
		<td class="header">订单管理</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a href="Order_info_List.asp" target="main">订单管理</a></li>
		   <li><a href="Order_info_search.asp" target="main">订单搜索</a></li>
		   <li><a href="Order_info_recycle.asp" target="main">订单还原</a></li>
		   <li><a href="Order_info_Print.asp" target="main">订单打印</a></li>
		   <li><a href="Order_info_SaleCount.asp" target="main">销售统计</a></li>
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
		<td class="header">会员管理</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a href="user_option_set.asp" target="main">会员选项</a></li>
		   <li><a href="user_info_list.asp" target="main">会员信息管理</a></li>
		   <li><a href="user_info_search.asp" target="main">会员高级搜索</a></li>
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
		<td class="header">文章管理</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a target="main" href="News_Info_Add.asp">添加文章</a></li>
		   <li><a href="news_info_list.asp" target="main">管理文章</a></li>
		   <li><a href="help_info_add.asp" target="main">帮助信息添加</a></li>
		   <li><a href="help_info_list.asp" target="main">帮助信息管理</a></li>
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
		<td class="header">留言管理</td>
	</tr>
	<tr>
		<td class="altbg2">
		    <li><a target="main" href="GB_Info_List.asp">在线留言管理</a></li>
  			<li><a target="main" href="Prod_Review_List.asp"> 商品评论管理</a></li>		   
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
		<td class="header">权限管理</td>
	</tr>
	<tr>
		<td class="altbg2">
		   <li><a href="admin_info_add.asp" target="main">管理人员添加</a></li>
		   <li><a href="admin_info_list.asp" target="main">管理人员管理</a></li>
		   <li><a href="admin_info_PassWordModiByUserName.asp?admin_info_UserName=<%=session("admin_info_UserName")%>" target="main">管理密码修改</a></li>
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
		<td align="center"><font color="#999999">电子商务管理系统<br>
		设计者：30818103</font></a></td>
	</tr>
</table>
</body>

</html>

