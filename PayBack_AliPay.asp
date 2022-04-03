<center><%dim dbpath
dbpath=""
%>
<!--#include file="conn.asp"-->
<!--#include file="include/MyRequest.asp"-->
<!--#include file="include/base64.asp"-->
<!--#include file="include/mmd5.asp"-->
<%
sql="select base_NetPay_AlipayEmail,base_NetPay_AlipaySafeCode from base_NetPay where base_NetPay_id=1"
set rs=conn.execute (sql)
base_NetPay_AlipayEmail   =rs(0)
base_NetPay_AlipaySafeCode=rs(1)
rs.close
set rs=nothing
constPayEmail        = replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipayEmail))),chr(13)&chr(10),"<br>")
constPaySecurityCode = replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipaySafeCode))),chr(13)&chr(10),"<br>")

msg_id        = my_request("msg_id",0)              '通知ID
order_no      = my_request("order_no",0)            '交易订单号
gross         = my_request("gross",0)               '交易总金额
buyer_email   = my_request("buyer_email",0)         '买家的支付宝账户
buyer_name    = my_request("buyer_name",0)          '买家姓名
buyer_address = my_request("buyer_address",0)       '买家地址
buyer_zipcode = my_request("buyer_zipcode",0)       '买家邮编
buyer_tel     = my_request("buyer_tel",0)           '买家电话号码
buyer_mobile  = my_request("buyer_mobile",0)        '买家手机号码
action        = my_request("action",0)              '通知动作
Req_Date      = my_request("date",0)                '发送通知时的支付宝系统当前时间，格式为：yyyyMMddHHmmss
ac            = my_request("ac",0)

'###############################
'   检查信息是否由支付宝发出
'###############################
if action <> "test" then        
    Set Retrieval  = Server.CreateObject("Microsoft.XMLHTTP") 
    strURL = "http://notify.alipay.com/trade/notify_query.do?msg_id=" & msg_id
    strURL = strURL & "&email=" & constPayEmail & "&order_no=" & order_no
    Retrieval.open "GET", strURL, False, "", "" 
    Retrieval.send()
    ReturnState = Retrieval.ResponseText
    Set Retrieval = Nothing
    conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','"&strURL&" 来源状态:"&ReturnState&"')"
    If Cstr(ReturnState) <> "true" and Cstr(ReturnState) <> "false" Then
        conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','"&strURL&" 来源错误完成')"
    End If    
end if

Select Case action
    Case "test"
        response.write "Y"
        conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','测试接口')"
        
    Case "sendOff"        '用户已付款

        Str = "msg_id" & msg_id & "order_no" & order_no & "gross" & gross  & "buyer_email" & buyer_email & "buyer_name" & buyer_name & "buyer_address" & buyer_address & "buyer_zipcode" & buyer_zipcode & "buyer_tel" & buyer_tel & "buyer_mobile" & buyer_mobile & "action" & action  & "date" & Req_Date & constPaySecurityCode

        if MD5(Str) = ac then                        
            conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('"&order_no&"','"&ac&" 通过')"
            conn.execute("Update buyer set zt =1 where ddbh='"&order_no&"'")
            response.write "Y"
        else
            conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('"&order_no&"','"&Str&"-"&ac&" AC不通过')"
            response.write "N"
        end if
                
    Case "checkOut"        '交易完成
         conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('"&order_no&"','"&Str&"-"&ac&" 交易完成')"
         conn.execute("Update buyer set zt =3 where ddbh='"&order_no&"'")
         response.write "Y"
         
    Case Else
         conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','其他情况')"                
         response.write "N"
                
End Select

conn.close
set conn=nothing
%>


 
</center>