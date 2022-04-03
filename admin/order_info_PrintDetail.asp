<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=2
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
response.buffer = true

Order_info_no=my_request("Order_info_no",0)
sql="select * from Order_info where Order_info_no='"&Order_info_no&"'"
set rs=conn.execute (sql)
if rs.eof then
    response.write "<SCRIPT language=JavaScript>alert('出错了，没有此订单。');"
    response.write "location.href='javascript:history.go(-1)';</SCRIPT>"
    response.end
else
    order_info_no           =rs("order_info_no")
    order_info_RealName     =rs("order_info_RealName")
    order_info_mobile       =rs("order_info_mobile")
    order_info_tel          =rs("order_info_tel")
    order_info_address      =rs("order_info_address")
    order_info_zip          =rs("order_info_zip")
    order_info_pay          =rs("order_info_pay")
    order_info_deliver      =rs("order_info_deliver")
    order_info_DeliverCost  =rs("order_info_DeliverCost")
    order_info_ProdCost     =rs("order_info_ProdCost")
    order_info_AllCost      =rs("order_info_AllCost")
    order_info_BuyNote      =rs("order_info_BuyNote")
    order_info_BuyTime      =rs("order_info_BuyTime")
    order_info_BuyIP        =rs("order_info_BuyIP")
    order_info_ProdIds      =rs("order_info_ProdIds")
    order_info_ProdNums     =rs("order_info_ProdNums")
    order_info_ProdPrices   =rs("order_info_ProdPrices")
    order_info_ProdNames    =rs("order_info_ProdNames")
    order_info_uid          =rs("order_info_uid")
    order_info_UserName     =rs("order_info_UserName")
    order_info_CheckStates  =rs("order_info_CheckStates")
    order_info_CheckNote    =rs("order_info_CheckNote")
    order_info_CheckTime    =rs("order_info_CheckTime")
    order_info_CheckRealName=rs("order_info_CheckRealName")
end if
rs.close
set rs=nothing

select case order_info_CheckStates
    case 0
        order_info_CheckStates="新订单(未确认)"
    case 1
        order_info_CheckStates="会员自行取消"
    case 2
        order_info_CheckStates="无效单，已取消"
    case 3
        order_info_CheckStates="已确认，待付款"
    case 4
        order_info_CheckStates="已发货，待收货"
    case 5
        order_info_CheckStates="在线支付成功"
    case 6
        rder_info_CheckStates="订单完成"
end select 

'告诉浏览器用word来显示文档 
Response.ContentType = "application/msword" 
'文档设定 
response.AddHeader "content-disposition", "inline; filename="&Order_info_no&"号订单.doc" 
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body>
<h2><%=Order_info_No%>号订单信息:</h2>
<p>订单时间：<br>
订单金额：￥<%=Order_info_AllCost%><br>
订单号：<%=Order_info_No%><br>
会员ID：<%=Order_info_UserName%><br>
配送方式：<%=Order_info_Deliver%><br>
收货人姓名：<%=Order_info_RealName%><br>
收货地址：<%=Order_info_address%>&nbsp;&nbsp;(邮编:<%=Order_info_zip%>)<br>
联系电话：<%=Order_info_Tel%><br>
手机：<%=Order_info_Mobile%><br>
电子邮箱：<%=Order_info_email%><br>
顾客说明：<%=Order_info_BuyNote%><br>
订单处理说明：<%=Order_info_CheckNote%><br>
订单状态：<%=Order_info_CheckStates%><br>
　</p>
<h2><b>购物清单:</b></h2>
<table border="1" width="100%" style="border-collapse: collapse" bordercolor="#000000" cellspacing="0" cellpadding="4">
					<tr>
						<td><b>商品名称</b></td>
						<td><b>购买数量</b></td>
						<td><b>结算单价</b></td>
						<td><b>小计</b></td>
					</tr>		
			        <%
                    a=split(order_info_ProdIds,",")
                    b=split(order_info_ProdNums,",")
                    c=split(order_info_ProdPrices,",")
                    d=split(order_info_ProdNames,",")
                    for i=0 to ubound(a)
                        YourID=a(i)
                        YourBuyNums=b(i)
                        YourPrice=c(i)
                        YourProdName=d(i)
                    %>
					<tr>
						<td><%=YourProdName%></td>
						<td><%=YourBuyNums%></td>
						<td>￥<%=YourPrice%></td>
						<td>￥<%=YourPrice*YourBuyNums%></td>
					</tr>
					<%next%>
					<tr>
					    <td colspan=4>
					    商品总价：￥<%=order_info_ProdCost%><br>
				        配送费用：￥<%=order_info_DeliverCost%><br>
				        总计金额：<font color="#FF0000"><b>￥<%=order_info_AllCost%></b></font>
					    </td>
					</tr>
				</table>
				

</body>

</html>

 
