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
  response.write("<script>alert(""��������!"");location.href=""order_info_List.asp"";</script>")
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
        order_info_pay="֧��������֧��"
    case 5
        order_info_pay="PayPal����֧��"
    case 6 
        order_info_pay="���л��"
    case 7
        order_info_pay="�ʾֻ��"
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
		
		'���������ɣ������Ӷ�����Ϣ����ɵĶ��������嵥��
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
            	//��ȥ�����
            	conn.execute ("update [product_info] set product_info_kucun=product_info_kucun-"&YourBuyNum&" where id="&YourID)
            next
        end if
        
        
        
        '���ԭ����״̬Ϊ���״̬�ұ��δ����޸�Ϊδ���״̬�������ӵ������嵥���Ķ�����ϢӦ������
        if OldStates=6 and order_info_CheckStates<>6 then
            conn.execute ("delete from [order_buy] where order_buy_InfoId="&id)
        end if
        call ok("���ѳɹ�������һ��������Ϣ��","order_info_list.asp")
    end if
end sub
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ-�鿴/�༭</title>
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
		<td colspan="2" class="header">������Ϣ�鿴/�༭</td>
	</tr>
	<tr>
		<td>�µ�ʱ�䣺</td>
		<td><%=order_info_BuyTime%></td>
	</tr>
<tr>
		<td>�����ţ�</td>
		<td><%=order_info_no%></td>
	</tr>
<tr>
		<td>������</td>
		<td><%=order_info_AllCost%>Ԫ</td>
	</tr>
<tr>
		<td>��Ա�û�����</td>
		<td><a href=user_info_modi.asp?user_info_id=<%=order_info_uid%>><%=order_info_UserName%></a></td>
	</tr>
<tr>
		<td>���ͷ�ʽ��</td>
		<td><%=order_info_deliver%></td>
	</tr>
<tr>
		<td>���ʽ��</td>
		<td><%=order_info_pay%></td>
	</tr>
<tr>
		<td>�ջ���������</td>
		<td><%=order_info_RealName%></td>
	</tr>
<tr>
		<td>��ϵ�绰��</td>
		<td><%=order_info_mobile%></td>
	</tr>
		<td>Email��</td>
		<td><%=order_info_email%></td>
	</tr>
<tr>
		<td>�ջ���ϸ��ַ��</td>
		<td><%=order_info_address%></td>
	</tr>
<tr>
		<td>�������룺</td>
		<td><%=order_info_zip%></td>
	</tr>
<tr>
		<td>�˿͸��ԣ�</td>
		<td><%=order_info_BuyNote%></td>
	</tr>
<tr>
		<td>��������˵����</td>
		<td><textarea rows="4" name="order_info_CheckNote" cols="60"><%=order_info_CheckNote%></textarea></td>
	</tr>
	<tr>
		<td>����״̬��</td>
		<td><select name="order_info_CheckStates">
		<option value="0" <%if order_info_CheckStates=0 then response.write "selected"%>>�¶���</option>
		<option value="1" <%if order_info_CheckStates=1 then response.write "selected"%>>��Ա����ȡ��</option>
		<option value="2" <%if order_info_CheckStates=2 then response.write "selected"%>>��Ч������ȡ��</option>
		<option value="3" <%if order_info_CheckStates=3 then response.write "selected"%>>��ȷ�ϣ�������</option>
		<option value="4" <%if order_info_CheckStates=4 then response.write "selected"%>>�ѷ��������ջ�</option>
		<option value="5" <%if order_info_CheckStates=5 then response.write "selected"%>>����֧���ɹ�</option>
		<option value="6" <%if order_info_CheckStates=6 then response.write "selected"%> >�������</option>
		</select> <%if order_info_CheckStates<>0 then%>( <b>����ʱ��</b>��<%=order_info_CheckTime%>  )<%end if%> 
		</td>
	</tr>
	<tr>
		<td>�����嵥��</td>
		<td>
		<table border="1" width="100%" style="border-collapse: collapse" bordercolor="#CCCCCC" cellspacing="0" cellpadding="4">
					<tr>
						<td bgcolor="#EAEAEA"><b>��Ʒ����</b></td>
						<td bgcolor="#EAEAEA"><b>��������</b></td>
						<td bgcolor="#EAEAEA"><b>���㵥��</b></td>
						<td bgcolor="#EAEAEA"><b>С��</b></td>
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
                        response.write "<td>��"&YouPrice&"</td>"
                        response.write "<td>��"&YouPrice*YouBuyNums&"</td></tr>"
                    next
                    
                    %>
				</table>
				�ϼ���Ʒ�۸�<font color="#FF0000"><b>��<%=order_info_ProdCost%></b></font><br>
				�˷ѣ�<font color="#FF0000"><b>��<%=order_info_DeliverCost%></b></font> (<%=order_info_deliver%>)<br>
				�ܼƣ�<font color="#FF0000"><b>��<%=order_info_AllCost%></b></font>
		</td>
	</tr>
	<tr>
		<td>��</td>
		<td>
		<input type="submit" value="�ύ" name="B1">&nbsp;&nbsp;&nbsp;
		<input type="reset" value="����" name="B2"></td>
	</tr>
</form>
</tbody>
</table>

</body>

</html>