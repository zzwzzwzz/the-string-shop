<center><%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
order_no=my_request("search_order_info_no",0)
set rs=server.createobject("adodb.recordset")
sql="select order_info_CheckStates,order_info_CheckTime,order_info_BuyTime from order_info where order_info_no='"&order_no&"'"
rs.open sql,conn,1,1
if rs.eof then 
    CheckStates="没有这个订单号,请确认您输入的订单号是否有误!"
else
    order_info_CheckStates  =rs(0)
    order_info_CheckTime    =rs(1)
    order_info_BuyTime      =rs(2)
end if
rs.close
set rs=nothing

select case order_info_CheckStates
    case "0"
        CheckStates="新订单(未确认)"
        order_info_CheckTime=order_info_BuyTime
    case "1"
        CheckStates="顾客自行取消"
    case "2"
        CheckStates="无效单，已取消"
    case "3"
        CheckStates="已确认，待付款"
    case "4"
        CheckStates="已发货，待收货"
    case "5"
        CheckStates="在线支付成功"
    case "6"
        CheckStates="订单完成"
end select

call up("订单查询结果","订单查询结果","订单查询结果")

response.write  "<tr><td>订单号： "&order_no&"</td></tr>"&_
				"<tr><td>处理状态： <font color=red>" &CheckStates&"</font></td></tr>"&_
				"<tr><td>处理时间： "&order_info_CheckTime&"</td></tr>"
call down()
%></center>