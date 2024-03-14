<%dim dbpath
dbpath="../"
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
sql="select count(order_info_id) as num0 from order_info where order_info_CheckStates=0 and order_info_recycle=0"
set rs=conn.execute (sql)
num0=rs("num0")
rs.close
set rs=nothing		
		
sql="select count(order_info_id) as num1 from order_info where order_info_CheckStates=1 and order_info_recycle=0"
set rs=conn.execute (sql)
num1=rs("num1")
rs.close
set rs=nothing	
		
sql="select count(order_info_id) as num2 from order_info where order_info_CheckStates=2 and order_info_recycle=0"
set rs=conn.execute (sql)
num2=rs("num2")
rs.close
set rs=nothing	

sql="select count(order_info_id) as num3 from order_info where order_info_CheckStates=3 and order_info_recycle=0"
set rs=conn.execute (sql)
num3=rs("num3")
rs.close
set rs=nothing		
		
sql="select count(order_info_id) as num4 from order_info where order_info_CheckStates=4 and order_info_recycle=0"
set rs=conn.execute (sql)
num4=rs("num4")
rs.close
set rs=nothing	
		
sql="select count(order_info_id) as num5 from order_info where order_info_CheckStates=5 and order_info_recycle=0"
set rs=conn.execute (sql)
num5=rs("num5")
rs.close
set rs=nothing	

sql="select count(order_info_id) as num6 from order_info where order_info_CheckStates=6 and order_info_recycle=0"
set rs=conn.execute (sql)
num6=rs("num6")
rs.close
set rs=nothing

sql="select sum(order_buy_ProdPrice*order_buy_ProdNum) as sumsell from order_buy"
set rs=conn.execute (sql)
sumsell=rs("sumsell")
rs.close
set rs=nothing

sql="select sum(order_buy_ProdNum) as sumnum from order_buy"
set rs=conn.execute (sql)
sumnum=rs("sumnum")
rs.close
set rs=nothing

sql="select count(id) as pnum from product_info"
set rs=conn.execute (sql)
pnum=rs("pnum")
rs.close
set rs=nothing

sql="select count(prod_BigClass_id) as bnum from prod_BigClass"
set rs=conn.execute (sql)
bnum=rs("bnum")
rs.close
set rs=nothing

sql="select count(prod_SmallClass_id) as snum from prod_SmallClass"
set rs=conn.execute (sql)
snum=rs("snum")
rs.close
set rs=nothing

sql="select count(prod_review_id) as prnum from prod_review"
set rs=conn.execute (sql)
prnum=rs("prnum")
rs.close
set rs=nothing

sql="select count(guest_info_id) as gnum from guest_info"
set rs=conn.execute (sql)
gnum=rs("gnum")
rs.close
set rs=nothing

sql="select count(user_info_id) as unum from user_info"
set rs=conn.execute (sql)
unum=rs("unum")
rs.close
set rs=nothing
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台-首页</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
	<tr>
		<td class="header">后台首页</td>
	</tr>
	<tr>
		<td class="altbg2" colspan="6"></td>
	</tr>
	<tr>
		<td class="altbg1">消息提醒</td>
	</tr>
	<tr>
		<td>
		<p class="p2">
				<a href="order_info_list.asp">&nbsp;&nbsp;您有&nbsp;&nbsp;<%=num0%>&nbsp; 笔新订单等待处理！</a></td>
	</tr>
	<tr>
		<td class="altbg1">信息统计</td>
	</tr>
	<tr>
		<td>
		<table border="0" width="100%" id="table1" cellpadding="4" style="border-collapse: collapse">
			<tr>
				<td valign="top" style="border-bottom: 1px solid #E4E4E4"><b>订单信息：
		    </b>
		   </td>
				<td style="border-bottom: 1px solid #E4E4E4"> <li>会员自行取消订单&nbsp;&nbsp; ：&nbsp;&nbsp;<%=num1%>&nbsp; 笔</li>
		    <li>管理员取消订单&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ：&nbsp;&nbsp;<%=num2%>&nbsp; 笔</li>
		    <li>已确认，待付款订单：&nbsp;&nbsp;<%=num3%>&nbsp; 笔</li>
		    <li>已发货，待收货订单：&nbsp;&nbsp;<%=num4%>&nbsp; 笔</li>
		    <li>在线支付完成订单&nbsp;&nbsp; ：&nbsp;&nbsp;<%=num5%>&nbsp; 笔</li>
		    <li>已销售完成订单&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ：&nbsp;&nbsp;<%=num6%>&nbsp; 笔</td>
			</tr>
			<tr>
				<td style="border-bottom: 1px solid #E4E4E4">
			<b>销售状况：</b></td>
				<td style="border-bottom: 1px solid #E4E4E4"><li>共计已完成销售量：<%=sumnum%></li> 
				<li>销售额：<%=sumsell%> RMB</li>
			</td>
			</tr>
			<tr>
				<td style="border-bottom: 1px solid #E4E4E4"><b>商品统计： </b></td>
				<td style="border-bottom: 1px solid #E4E4E4"><li>大类别：<%=bnum%> </li> <li>小类别：<%=snum%> </li> <li>商品数量：<%=pnum%> </li>
				<li>商品评论：<%=prnum%> 条&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </li></td>
			</tr>
			<tr>
				<td style="border-bottom: 1px solid #E4E4E4"><b>留言信息：</b></td>
				<td style="border-bottom: 1px solid #E4E4E4"><%=gnum%> 条</td>
			</tr>
			<tr>
				<td><b>会员信息：</b></td>
				<td><%=unum%> 条</td>
			</tr>
		</table>
		</td>
	</tr>
</tbody>
</table>

</body>

</html>