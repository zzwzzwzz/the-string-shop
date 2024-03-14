<center>
<center><!--#include file="User_Chk.asp"-->
<%
dim dbpath
dbpath=""
%>
<!--#include file=Conn.asp -->
<!--#include file=include/MyRequest.asp -->
<!--#include file=Sub.asp -->
<%
id=my_request("id",1)
if id="" or isnull(id) or IsNumeric(id)=False then
  response.write("<script>alert(""��������!"");location.href=""user_orderList.asp"";</script>")
  response.end
end if

sql="select * from order_info where order_info_id="&id
set rs=conn.execute (sql)
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
order_info_CheckStates  =rs("order_info_CheckStates")
order_info_CheckTime    =rs("order_info_CheckTime")
rs.close
set rs=nothing

select case order_info_pay
    case 1
        order_info_pay="֧��������֧��"
    case 5
        order_info_pay="PayPal����֧��"
end select

select case order_info_CheckStates
    case 0
        order_info_CheckStatesTxt="�¶���(δȷ��)"
    case 1
        order_info_CheckStatesTxt="��Ա����ȡ��"
    case 2
        order_info_CheckStatesTxt="��Ч������ȡ��"
    case 3
        order_info_CheckStatesTxt="��ȷ�ϣ�������"
    case 4
        order_info_CheckStatesTxt="�ѷ��������ջ�"
    case 5
        order_info_CheckStatesTxt="����֧���ɹ�"
    case 6
        order_info_CheckStatesTxt="�������"
end select   

action=my_request("action",0)
if action="cancel" then
    call Cancel()
end if

'ȡ������
sub Cancel()
    response.write "<hr>"
  	id=my_request("id",1)
  	sql="update order_info set order_info_CheckStates=1 where order_info_id="&id
  	set rs=conn.execute (sql)
  	response.write "<script language=javascript>alert('���Ķ�����ȡ���ɹ���');location.href=""user_OrderList.asp"";</script>"
  	response.End
end sub

call up("������ϸ","������ϸ","������ϸ")
%>
<!--#include file="User_Menu.asp"-->
<%
response.write  "<tr><td colspan=2 align=center><b>������ϸ���</b> 		</td></tr>"&_
				"<tr><td>�������:  </td><td>"&order_info_no&"	   		</td></tr>"&_
				"<tr><td>�µ�ʱ��:  </td><td>"&order_info_BuyTime&" 		</td></tr>"&_
				"<tr><td>�������:  </td><td>��"&formatnumber(order_info_AllCost,2,-1)&"</td></tr>"&_
				"<tr><td>���ͷ�ʽ:  </td><td>"&order_info_Deliver&" 		</td></tr>"&_
				"<tr><td>���ʽ:  </td><td>"&order_info_Pay&"     		</td></tr>"&_
				"<tr><td>�ջ�������:</td><td>"&order_info_RealName&"		</td></tr>"&_
				"<tr><td>��ϵ�绰:  </td><td>"&order_info_Mobile&"     		</td></tr>"&_
				"<tr><td>Email:     </td><td>"&order_info_email&"		</td></tr>"&_
				"<tr><td>�ջ���ַ:  </td><td>"&order_info_address&"		</td></tr>"&_
				"<tr><td>��������:  </td><td>"&order_info_zip&"			</td></tr>"&_
				"<tr><td>�˿͸���:  </td><td>"&order_info_BuyNote&"		</td></tr>"&_
				"<tr><td>����״̬:  </td><td><b>"&order_info_CheckStatesTxt&"</b>"
									if order_info_CheckStates=0 then
										response.write "<input class=button type=button value=���ȡ���˶��� onclick=window.location='user_orderDetail.asp?id="&id&"&action=cancel'>"
									end if
response.write  "					</td></tr>"&_
				"<tr><td>�����嵥:  </td><td>"

'<!--cartlist begin-->
response.write  "	<table border=1 width=100% style='border-collapse: collapse' bordercolor=#CCCCCC cellspacing=0 cellpadding=4>"&_
				"		<tr><td><b>��Ʒ����</b></td><td><b>��������</b></td><td><b>���㵥��</b></td><td><b>С��</b></td></tr>"		
                    		a=split(order_info_ProdIds,",")
                    		b=split(order_info_ProdNums,",")
                    		c=split(order_info_ProdPrices,",")
                    		d=split(order_info_ProdNames,",")
                    		for i=0 to ubound(a)
                    			YourID=a(i)
                   				YourBuyNums=b(i)
                   				YourPrice=c(i)
                    			YourProdName=d(i)
response.write  "		<tr>"&_
				"		    <td><a href='product_detail.asp?id="&YourID&"' target=_blank>"&YourProdName&"</a></td>"&_
				"		    <td>"&YourBuyNums&"</td>"&_
				"		    <td>��"&YourPrice&"</td>"&_
				"		    <td>��"&YourPrice*YourBuyNums&"</td>"&_
				"		</tr>"
					        next
response.write  "	</table>"&_
				"	�ϼ���Ʒ�۸�<font color=#FF0000><b>��"&order_info_ProdCost&"</b></font><br>"&_
				"    �˷ѣ�<font color=#FF0000><b>��"&order_info_DeliverCost&"</b></font> ("&order_info_deliver&")<hr>"&_
				"    �ܼƣ�<font color=#FF0000><b>��"&order_info_AllCost&"</b></font>"
'<!--cartlist end-->

response.write "</td></tr>"
call down()
%></center>
</center>