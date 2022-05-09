<!--#include file="admin_check.asp"-->
<%dim dbpath,nownum
dbpath="../"
nownum=2
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_info_FlagCheck.asp"-->
<!--#include file="../include/MyRequest.asp"-->
<%
id=my_request("order_info_id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""参数错误!"");location.href=""order_info_List.asp"";</script>")
  response.end
end if

set rs=server.createobject("adodb.recordset")
sql="select * from order_info where order_info_id="&id
rs.open sql,conn,1,1
order_info_no           =rs("order_info_no")
order_info_RealName     =rs("order_info_RealName")
order_info_mobile       =rs("order_info_mobile")
order_info_email        =rs("order_info_email")
order_info_address      =rs("order_info_address")
order_info_zip          =rs("order_info_zip")
order_info_pay          =rs("order_info_pay")
order_info_deliver      =rs("order_info_deliver")
order_info_DeliverCost  =rs("order_info_DeliverCost")
order_info_ProdCost     =rs("order_info_ProdCost")
order_info_AllCost      =rs("order_info_AllCost")
order_info_BuyNote      =rs("order_info_BuyNote")
order_info_BuyTime      =rs("order_info_BuyTime")
order_info_ProdIds      =rs("order_info_ProdIds")
order_info_ProdNums     =rs("order_info_ProdNums")
order_info_ProdPrices   =rs("order_info_ProdPrices")
order_info_ProdNames    =rs("order_info_ProdNames")
order_info_uid          =rs("order_info_uid")
order_info_UserName     =rs("order_info_UserName")
order_info_CheckStates  =rs("order_info_CheckStates")
order_info_CheckNote    =rs("order_info_CheckNote")
order_info_CheckTime    =rs("order_info_CheckTime")
rs.close
set rs=nothing

select case order_info_pay
    case 1
        order_info_pay="支付宝在线支付"
    case 5
        order_info_pay="PayPal在线支付"
    case 6 
        order_info_pay="银行汇款"
    case 7
        order_info_pay="邮局汇款"
end select

action=my_request("action",0)
if action="save" then
    call save()
end if

sub save()
    id=my_request("order_info_id",1)
    OldStates=my_request("OldStates",1)
    order_info_CheckStates=my_request("order_info_CheckStates",1)
    order_info_CheckNote  =my_request("CheckNote",0)
    if id="" or order_info_CheckStates="" then
        call error()
    else
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from order_info where order_info_id="&id
        rs.open sql,conn,1,3
        rs("order_info_CheckStates")  =order_info_CheckStates
        rs("order_info_CheckNote")    =order_info_CheckNote
        rs("order_info_CheckTime")    =now()
        rs.update
        rs.close
        set rs=nothing
        
        if order_info_uid<>"" then
			Set rs=Server.CreateObject("ADODB.Recordset")
			sql="select root_option_MarkYuan from root_option where id=1"
			rs.open sql,conn,1,1
			root_option_MarkYuan=rs(0)
			rs.close
			set rs=nothing
			x=1/root_option_MarkYuan
			y=order_info_ProdCost/x
			y=cint(y)
		end if
		
		//如果订单完成，则添加订单信息到完成的订单购物清单表
        if OldStates<>6 and order_info_CheckStates=6 then
            sql="select order_info_BuyTime,order_info_ProdIds,order_info_ProdNums,order_info_ProdPrices,order_info_ProdNames from order_info where order_info_id="&id
            set rs=conn.execute (sql)
            order_info_BuyTime      =rs(0)
            order_info_ProdIds      =rs(1)
            order_info_ProdNums     =rs(2)
            order_info_ProdPrices   =rs(3)
            order_info_ProdNames    =rs(4)
            rs.close
            set rs=nothing
            
            a=split(order_info_ProdIds,",")
            b=split(order_info_ProdNums,",")
            c=split(order_info_ProdPrices,",")
            d=split(order_info_ProdNames,",")
            for i=0 to ubound(a)
                YourID=a(i)
                YourBuyNum=b(i)
                YourPrice=c(i)
                YourProdName=d(i)
                conn.execute ("insert into [order_buy] (order_buy_InfoId,order_buy_ProdId,order_buy_ProdNum,order_buy_ProdPrice,order_buy_ProdName,order_buy_BuyTime) values ("&id&","&YourID&","&YourBuyNum&","&YourPrice&",'"&YourProdName&"','"&order_info_BuyTime&"')")
            	//减去库存量
            	conn.execute ("update [product_info] set product_info_kucun=product_info_kucun-"&YourBuyNum&" where id="&YourID)
            next
        end if
        
        
        
        //如果原订单状态为完成状态且本次处理修改为未完成状态，则添加到购物清单表的订单信息应撤销掉
        if OldStates=6 and order_info_CheckStates<>6 then
            conn.execute ("delete from [order_buy] where order_buy_InfoId="&id)
        end if
        call ok("您已成功处理了一条订单信息！","order_info_list.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>订单信息-查看/编辑</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

<table cellspacing="1" cellpadding="4" width="100%" class="tableborder">
<tbody class="altbg2">
<form name="form1" action="order_info_Modi.asp" method="post">
<input type="hidden" name="action" value="save"> 
<input type="hidden" name="order_info_id" value="<%=id%>"> 
<input type="hidden" name="OldStates" value="<%=Order_info_CheckStates%>">
	<tr>
		<td colspan="2" class="header">订单信息查看/编辑</td>
	</tr>
	<tr>
		<td>下单时间：</td>
		<td><%=order_info_BuyTime%></td>
	</tr>
<tr>
		<td>订单号：</td>
		<td><%=order_info_no%></td>
	</tr>
<tr>
		<td>订单金额：</td>
		<td><%=order_info_AllCost%>元</td>
	</tr>
<tr>
		<td>会员用户名：</td>
		<td><a href=user_info_modi.asp?user_info_id=<%=order_info_uid%>><%=order_info_UserName%></a></td>
	</tr>
<tr>
		<td>配送方式：</td>
		<td><%=order_info_deliver%></td>
	</tr>
<tr>
		<td>付款方式：</td>
		<td><%=order_info_pay%></td>
	</tr>
<tr>
		<td>收货人姓名：</td>
		<td><%=order_info_RealName%></td>
	</tr>
<tr>
		<td>联系电话：</td>
		<td><%=order_info_mobile%></td>
	</tr>
		<td>Email：</td>
		<td><%=order_info_email%></td>
	</tr>
<tr>
		<td>收货详细地址：</td>
		<td><%=order_info_address%></td>
	</tr>
<tr>
		<td>邮政编码：</td>
		<td><%=order_info_zip%></td>
	</tr>
<tr>
		<td>顾客附言：</td>
		<td><%=order_info_BuyNote%></td>
	</tr>
<tr>
		<td>订单处理说明：</td>
		<td><textarea rows="4" name="order_info_CheckNote" cols="60"><%=order_info_CheckNote%></textarea></td>
	</tr>
	<tr>
		<td>订单状态：</td>
		<td><select name="order_info_CheckStates">
		<option value="0" <%if order_info_CheckStates=0 then response.write "selected"%>>新订单</option>
		<option value="1" <%if order_info_CheckStates=1 then response.write "selected"%>>会员自行取消</option>
		<option value="2" <%if order_info_CheckStates=2 then response.write "selected"%>>无效单，已取消</option>
		<option value="3" <%if order_info_CheckStates=3 then response.write "selected"%>>已确认，待付款</option>
		<option value="4" <%if order_info_CheckStates=4 then response.write "selected"%>>已发货，待收货</option>
		<option value="5" <%if order_info_CheckStates=5 then response.write "selected"%>>在线支付成功</option>
		<option value="6" <%if order_info_CheckStates=6 then response.write "selected"%> >订单完成</option>
		</select> <%if order_info_CheckStates<>0 then%>( <b>处理时间</b>：<%=order_info_CheckTime%>  )<%end if%> 
		</td>
	</tr>
	<tr>
		<td>购物清单：</td>
		<td>
		<table border="1" width="100%" style="border-collapse: collapse" bordercolor="#CCCCCC" cellspacing="0" cellpadding="4">
					<tr>
						<td bgcolor="#EAEAEA"><b>商品名称</b></td>
						<td bgcolor="#EAEAEA"><b>购买数量</b></td>
						<td bgcolor="#EAEAEA"><b>结算单价</b></td>
						<td bgcolor="#EAEAEA"><b>小计</b></td>
					</tr>		
			        <%
                    a1=split(order_info_ProdIds,",")
                    b1=split(order_info_ProdNums,",")
                    c1=split(order_info_ProdPrices,",")
                    d1=split(order_info_ProdNames,",")
                    e=ubound(a1)
                    
                    for v=0 to e
                        ttt=a1(v)
                        YouBuyNums=b1(v)
                        YouPrice=c1(v)
                        YouProdName=d1(v)
                        response.write "<tr><td><a target=_blank href='../product_detail.asp?id="
                        response.write ttt&"' target=_blank>"
                        response.write YouProdName&"</a></td>"

                        response.write "<td>"&YouBuyNums&"</td>"
                        response.write "<td>￥"&YouPrice&"</td>"
                        response.write "<td>￥"&YouPrice*YouBuyNums&"</td></tr>"
                    next
                    
                    %>
				</table>
				合计商品价格：<font color="#FF0000"><b>￥<%=order_info_ProdCost%></b></font><br>
				运费：<font color="#FF0000"><b>￥<%=order_info_DeliverCost%></b></font> (<%=order_info_deliver%>)<br>
				总计：<font color="#FF0000"><b>￥<%=order_info_AllCost%></b></font>
		</td>
	</tr>
	<tr>
		<td>　</td>
		<td>
		<input type="submit" value="提交" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="reset" value="重置" name="B2"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>

