<center>
<%dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<%dim nowplace
nowplace="add_order"
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select root_option_GuestOrderOnOff from root_option where id=1"
rs.open sql,conn,1,1
root_option_GuestOrderOnOff=rs(0)
rs.close
set rs=nothing
if root_option_GuestOrderOnOff=1 then
%>
<!--#include file="User_Chk.asp"-->
<%
end if
%>
<!--#include file="include/MyRequest.asp" -->
<!--#include file="include/nosql.asp" -->
<!--#include file=Sub.asp -->
<%
ProdIds  = Session("ProdIds")
ProdNums = Session("ProdNums")
Sum      = Session("sum")
sum=FormatNumber(sum,2,-1)
sum=cint(sum)

//会员信息调出
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select user_info_RealName,user_info_mobile,user_info_tel,user_info_address,user_info_zip,user_info_email from user_info where user_info_id="&session("user_info_id")
rs.open sql,conn,1,1
user_info_RealName=rs(0)
user_info_mobile  =rs(1)
user_info_tel     =rs(2)
user_info_address =rs(3)
user_info_zip     =rs(4)
user_info_email   =rs(5)
rs.close
set rs=nothing

call up("结算下订单","结算下订单","<a href=cart_list.asp>购物车</a> &raquo; 结算下订单")

response.write  "		<table border=1 width=100% cellpadding=4 style='border-collapse: collapse' bordercolor=#DFDFDF>"&_
				"			<tr><td>商品名称</td><td>市场价</td><td>网站价</td><td>订购数量</td><td>小计</td></tr>"
							if ProdIds<>"" then
								aaa=split(ProdNums,",")
								bbb=split(ProdIds,",")

								for i=0 to ubound(bbb)
        						set rs=server.createobject("adodb.recordset")
        						sql="select id,product_info_PriceM,product_info_PriceS,product_info_name from product_info where id="&bbb(i)
        						rs.open sql,conn,1,1
        						if rs.eof then
        							response.write  "<tr><td colspan=5 align=center><a href=index.asp>购物车为空，请返回选购商品</a></td></tr>"&_
        											"</table>"
        							response.end
        						else
        	    				    //如果是会员则根据积分享受折扣优惠价
                    				if session("user_info_id")<>"" then
                    					//获取会员积分值
                      					ssql="select user_info_mark from user_info where user_info_id="&session("user_info_id")
					  					set rss=conn.execute (ssql)
					  					point=rss(0)
					  					rss.close
					  					set rss=nothing
					
					  					//获取享受会员折率
					  					sql1="select user_level_rebate from user_level where user_level_markmin<="&point&" and user_level_markmax>="&point&""
					  					set rs1=conn.execute (sql1)
					  					rebate=Cdbl(rs1(0))
					  					RMB=rs(2)*rebate/100
					  					rs1.close
					  					set rs1=nothing
									else
					 					//不是会员不享受折扣优惠价
                      					RMB=rs(2)
									end if

            						set id=rs(0)
            						set product_info_PriceM=rs(1)
            						set product_info_name=rs(3)
            						While Not rs.EOF
            						x=aaa(i)
            						if aaa(i)="" then x=1
            						sum1=sum1 + csng(rmb) * x
            						sum=FormatNumber(sum1,2,-1)
response.write  "			<tr>"&_
				"				<td><a href=Product_Detail.asp?id="&id&" target=_blank>"&product_info_name&"</a></td>"&_
				"				<td>￥"&FormatNumber(product_info_PriceM,2,-1)&"</td>"&_
				"				<td><font color=#FF0000>￥"&FormatNumber(Rmb,2,-1)&"</font></td>"&_
				"				<td>"&x&"</td>"&_
				"				<td>￥"&FormatNumber((csng(rmb)*x),2,-1)&"</td>"&_
				"			</tr>"
						    		rs.MoveNext
    			    				Wend
    							end if
    							rs.close
    							set rs=nothing
    							next
response.write  "			<tr>"&_
				"				<td colspan=5 align=right>合计金额：<span style='color:#FF6633;font-size:18px;'>￥"&sum&"</span></td>"&_
				"			</tr>"
    						else
    							response.write "<tr><td colspan=5 align=center><a href=index.asp>购物车为空，请返回选购商品!</a></td></tr>"
    						end if
response.write  "		</table>"&_
				"		<br>"&_
				"		<table border=0 width=100% cellpadding=4 style=border-collapse: collapse>"&_
				"		<form name=form1 action=Cart_OrderSave.asp method=post onsubmit=return check_form();>"&_
				"			<tr><td colspan=2><b>收货人信息</b></td></tr>"&_
				"			<tr><td>姓名：    </td><td><input type=text name=Order_info_RealName size=30 value="&User_info_RealName&"></td></tr>"&_
				"			<tr><td>电子邮件：</td><td><input type=text name=order_info_email size=30 value="&User_info_email&">(必须含@)</td></tr>"&_
				"			<tr><td>收货地址：</td><td><input type=text name=order_info_address size=30 value="&User_info_address&"> </td></tr>"&_
				"			<tr><td>邮政编码：</td><td><input type=text name=order_info_zip size=30 value="&User_info_zip&">(6位数字)</td></tr>"&_
				"			<tr><td>联系电话：</td><td><input type=text name=order_info_mobile size=30 value="&User_info_mobile&">(11位数字)</td></tr>"&_
				"			<tr><td></td><td><input type=checkbox name=order_info_up value=1>用上述联系方法覆盖帐户信息</td></tr>"&_
				"			<tr><td>客户留言：</td><td><textarea rows=3 name=order_info_BuyNote cols=50></textarea></td></tr>"&_
				"			<tr><td colspan=2><b>送货方式</b></td></tr>"&_
				"			<tr><td> </td>"&_
				"				<td>"
         						set rs=server.createobject("adodb.recordset")
            					sql="select root_deliver_name,root_deliver_cost,root_deliver_day from root_deliver order by id desc"
            					rs.open sql,conn,1,1
            					if not rs.eof then 
                					set root_deliver_name=rs(0)
                					set root_deliver_cost=rs(1)
                					set root_deliver_day =rs(2)
                					while not rs.eof
                					response.write "<input type=radio value="&root_deliver_name&" name=order_info_deliver>"&root_deliver_name&"  ( 费用："&formatnumber(root_deliver_cost,2,-1)&"元    时间："&root_deliver_day&"天 ) <br>"
                					rs.movenext
                					wend
            					end if
            					rs.close
            					set rs=nothing
response.write  "				</td>"&_
				"			</tr>"&_
				"			<tr><td colspan=2><b>付款方式</b></td></tr>"&_
				"			<tr><td> </td>"&_
				"				<td>"
								Set rs=Server.CreateObject("ADODB.Recordset")
					    		sql="select base_NetPay_AlipayOnOff,base_NetPay_ChinaBankOnOff,base_NetPay_IpayOnOff,base_NetPay_NpsOnOff,base_NetPay_PayPalOnOff from root_NetPay where base_NetPay_id=1"
					    		rs.open sql,conn,1,1
					    		base_NetPay_AlipayOnOff        =rs(0)
					    		base_NetPay_ChinaBankOnOff     =rs(1)
					    		base_NetPay_IpayOnOff          =rs(2)
					    		base_NetPay_NpsOnOff           =rs(3)
					    		base_NetPay_PayPalOnOff        =rs(4)
					    		rs.close
					    		set rs=nothing
					    
					    		if base_NetPay_AlipayOnOff=0 then
                            		response.write "<input type=radio value=1 name=order_info_pay>支付宝<img src=images/netpaylogo/NetPay_logo_alipay.gif align=absmiddle><br>"
                        		end if
					    		if base_NetPay_ChinaBankOnOff=0 then
                            		response.write "<input type=radio value=2 name=order_info_pay>网银在线<img src=images/netpaylogo/NetPay_logo_chinabank.gif align=absmiddle><br>"
                        		end if
					    		if base_NetPay_IpayOnOff=0 then
                            		response.write "<input type=radio value=3 name=order_info_pay>IPAY<img src=images/netpaylogo/NetPay_logo_ipay.gif align=absmiddle><br>"
                        		end if
					    		if base_NetPay_NpsOnOff=0 then
                            		response.write "<input type=radio value=4 name=order_info_pay>NPS<img src=images/netpaylogo/NetPay_logo_nps.gif align=absmiddle><br>"
                        		end if
					    		if base_NetPay_PayPalOnOff=0 then
                            		response.write "<input type=radio value=5 name=order_info_pay>贝宝PayPal<img src=images/netpaylogo/NetPay_logo_paypal.gif align=absmiddle><br>"
                        		end if
response.write  "			    <input type=radio value=6 name=order_info_pay>银行汇款<br>"&_
				"			    <input type=radio value=7 name=order_info_pay>邮局汇款"&_
				"				</td>"&_
				"			</tr>"&_
				"			<tr><td>　</td><td><input class=button type=submit value=  提交订单   ></td></tr>"&_
				"		</form>" &_ 
				"		</table>" &_  
				"</td></tr>"
call down()
%>
</center>