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

msg_id        = my_request("msg_id",0)              '֪ͨID
order_no      = my_request("order_no",0)            '���׶�����
gross         = my_request("gross",0)               '�����ܽ��
buyer_email   = my_request("buyer_email",0)         '��ҵ�֧�����˻�
buyer_name    = my_request("buyer_name",0)          '�������
buyer_address = my_request("buyer_address",0)       '��ҵ�ַ
buyer_zipcode = my_request("buyer_zipcode",0)       '����ʱ�
buyer_tel     = my_request("buyer_tel",0)           '��ҵ绰����
buyer_mobile  = my_request("buyer_mobile",0)        '����ֻ�����
action        = my_request("action",0)              '֪ͨ����
Req_Date      = my_request("date",0)                '����֪ͨʱ��֧����ϵͳ��ǰʱ�䣬��ʽΪ��yyyyMMddHHmmss
ac            = my_request("ac",0)

' �����Ϣ�Ƿ���֧��������
if action <> "test" then        
    Set Retrieval  = Server.CreateObject("Microsoft.XMLHTTP") 
    strURL = "http://notify.alipay.com/trade/notify_query.do?msg_id=" & msg_id
    strURL = strURL & "&email=" & constPayEmail & "&order_no=" & order_no
    Retrieval.open "GET", strURL, False, "", "" 
    Retrieval.send()
    ReturnState = Retrieval.ResponseText
    Set Retrieval = Nothing
    conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','"&strURL&" ��Դ״̬:"&ReturnState&"')"
    If Cstr(ReturnState) <> "true" and Cstr(ReturnState) <> "false" Then
        conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','"&strURL&" ��Դ�������')"
    End If    
end if

Select Case action
    Case "test"
        response.write "Y"
        conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','���Խӿ�')"
        
    Case "sendOff"        '�û��Ѹ���

        Str = "msg_id" & msg_id & "order_no" & order_no & "gross" & gross  & "buyer_email" & buyer_email & "buyer_name" & buyer_name & "buyer_address" & buyer_address & "buyer_zipcode" & buyer_zipcode & "buyer_tel" & buyer_tel & "buyer_mobile" & buyer_mobile & "action" & action  & "date" & Req_Date & constPaySecurityCode

        if MD5(Str) = ac then                        
            conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('"&order_no&"','"&ac&" ͨ��')"
            conn.execute("Update buyer set zt =1 where ddbh='"&order_no&"'")
            response.write "Y"
        else
            conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('"&order_no&"','"&Str&"-"&ac&" AC��ͨ��')"
            response.write "N"
        end if
                
    Case "checkOut"        '�������
         conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('"&order_no&"','"&Str&"-"&ac&" �������')"
         conn.execute("Update buyer set zt =3 where ddbh='"&order_no&"'")
         response.write "Y"
         
    Case Else
         conn.Execute"INSERT INTO [pay_back] (order_no,pay_value) VALUES ('0','�������')"                
         response.write "N"
                
End Select

conn.close
set conn=nothing
%>


 
</center>