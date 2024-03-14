<center>
<%
dim dbpath
dbpath=""
%>
<!--#include file="Conn.asp"-->
<!--#include file="include/MyRequest.asp" -->
<!--#include file="include/base64.asp"--> 
<!--#include file="alipay/alipay_md5.asp"--> 
<!--#include file="alipay/alipay_payto.asp"--> 
<!--#include file="include/pay.asp"-->
<%
Sub PutToShopBag( x, y )
    If Len(y) = 0 Then
        y = x
    ElseIf InStr(y, x ) <= 0 Then
        y = y & ","& x
    End If
End Sub
   
uid=session("user_info_id")
UserName=session("user_info_UserName")
ProdIds=Session("ProdIds")
ProdNums=Session("ProdNums")
   
if ProdIds<>"" then
    Products = Split(ProdIds,",")
    For i=0 To UBound(Products)
        sql="select Product_info_name,product_info_PriceS from product_info where id="&Products(i)
        Set rs=conn.Execute( sql )
        Product_info_name   = rs(0)
        product_info_PriceS = rs(1)
        PutToShopBag Product_info_name, ProdNames
        PutToShopBag product_info_PriceS, ProdPrices
    Next
end if

'���ﳵ״̬�ж�
If Len(ProdIds) = 0 Then
    response.write "<script language=javascript>alert('�Բ������Ĺ��ﳵΪ�գ�');location.href=""index.asp"";</script>"
    response.End
end if

'�������ڣ���ʽ��YYYYMMDD
yy=year(date)
mm=right("00"&month(date),2)
dd=right("00"&day(date),2)
'���ɶ�������������Ԫ��,��ʽΪ��Сʱ�����ӣ���
xiaoshi=right("00"&hour(time),2)
fenzhong=right("00"&minute(time),2)
miao=right("00"&second(time),2)
order_info_no         =yy & mm & dd & xiaoshi & fenzhong & miao
  
order_info_RealName   =my_request("order_info_RealName",0)
order_info_mobile     =my_request("order_info_mobile",0)
order_info_email	  =my_request("order_info_email",0)
order_info_address    =my_request("order_info_address",0)
order_info_zip        =my_request("order_info_zip",0)
order_info_pay        =my_request("order_info_pay",0)
order_info_ProdCost   =session("sum")
order_info_BuyNote    =my_request("order_info_BuyNote",0)
order_info_pay        =my_request("order_info_pay",0)
order_info_deliver    =my_request("order_info_deliver",0)
order_info_up         =my_request("order_info_up",1)
  
Set rs= Server.CreateObject("ADODB.Recordset")
sql="select root_deliver_cost from root_deliver where root_deliver_name='"&order_info_deliver&"'"
rs.open sql,conn,1,1
order_info_DeliverCost=rs(0)
rs.close
set rs=nothing
   
order_info_AllCost=order_info_DeliverCost+order_info_ProdCost
order_info_AllCost=FormatNumber(order_info_AllCost,2,-1)


if order_info_no="" or order_info_RealName="" or (order_info_mobile="") or order_info_address="" or order_info_pay="" or order_info_deliver="" then
    response.write "<script language='javascript'>"
    response.write "alert('��Ϣ��д��������');"
    response.write "location.href='javascript:history.go(-1)';"
    response.write "</script>"
    response.end
else
    Set rs= Server.CreateObject("ADODB.Recordset")
    sql="select * from Order_info"
    rs.open sql,conn,1,3
    rs.addnew
    rs("order_info_no")         =order_info_no
    rs("order_info_RealName")   =order_info_RealName
    rs("order_info_mobile")     =order_info_mobile
    rs("order_info_email")      =order_info_email
    rs("order_info_address")    =order_info_address
    rs("order_info_zip")        =order_info_zip
    rs("order_info_pay")        =order_info_pay
    rs("order_info_deliver")    =order_info_deliver
    rs("order_info_DeliverCost")=order_info_DeliverCost
    rs("order_info_ProdCost")   =order_info_ProdCost
    rs("order_info_AllCost")    =order_info_AllCost
    rs("order_info_BuyNote")    =order_info_BuyNote
    rs("order_info_BuyTime")    =now()
    rs("order_info_ProdIds")    =ProdIds    '����������Ʒid����
    rs("order_info_ProdNums")   =ProdNums   '����������Ʒ��������
    rs("order_info_ProdNames")  =ProdNames  '����������Ʒ���Ƽ���
    rs("order_info_ProdPrices") =ProdPrices '����������Ʒ���ۼ���
    rs("order_info_uid")        =uid
    rs("order_info_UserName")   =UserName      
    rs.update
    rs.close
    set rs=nothing 
    
    // �Ƿ���������ϵ���������ʻ���Ϣ  
    if order_info_up=1 then
        Set rs= Server.CreateObject("ADODB.Recordset")
        sql="select * from user_info where user_info_id="&uid
        rs.open sql,conn,1,3
        rs("user_info_RealName")=order_info_RealName
        rs("user_info_mobile")  =order_info_mobile
        rs("user_info_address") =order_info_address
        rs("user_info_zip")     =order_info_zip
        rs.update
        rs.close
        set rs=nothing
    end if
       
    'ɾ�����ﳵ����Ʒid�������ĻỰֵ
    Session.Contents.Remove("ProdIds")
    Session.Contents.Remove("ProdNums")
    Session.Contents.Remove("ProdPrices")
    'Session.Contents.Remove("ProdNames")
    Session.Contents.Remove("sum")
    Session.Contents.Remove("y")
       
%>
<script LANGUAGE="JavaScript"> 
window.open ("Cart_OrderOk.asp?Order_info_no=<%=order_info_no%>&order_info_AllCost=<%=order_info_AllCost%>", "newwindow", "height=250, width=400, top=0, left=0,toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no") 
</script>
<%        //���ʽ����
       select case order_info_pay
           case 1  '֧����
               sql="select base_NetPay_AlipayEmail,base_NetPay_AlipaySafeCode,base_NetPay_AlipayPartnerID from root_NetPay where base_NetPay_id=1"
               set rs=conn.execute (sql)
               base_NetPay_AlipayEmail   =rs(0)
               base_NetPay_AlipaySafeCode=rs(1)
               base_NetPay_AlipayPartnerID=rs(2)
               rs.close
               set rs=nothing
               v_mid=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipayEmail))),chr(13)&chr(10),"<br>")
               key1=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipaySafeCode))),chr(13)&chr(10),"<br>")
               partnerID=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipayPartnerID))),chr(13)&chr(10),"<br>")
               
               call pay_AliPay()
               Session.Contents.Remove("ProdNames")
         case 5  'PayPal
             sql="select root_NetPay_PayPalEmail from base_NetPay where root_NetPay_id=1"

             set rs=conn.execute (sql)
             base_NetPay_PayPalEmail  =rs(0)
             rs.close
             set rs=nothing 
             v_mid=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_PayPalEmail))),chr(13)&chr(10),"<br>")
             call pay_paypal()
                   
        case else
            response.redirect "Cart_OrderOk.asp?Order_info_no="&order_info_no&"&order_info_AllCost="&order_info_AllCost&""
            response.end      
    end select
end if                            
%>
</center>