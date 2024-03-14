<center><%dim dbpath
dbpath=""
%>
<!--#include file="conn.asp"-->
<!--#include file="include/base64.asp"-->
<!--#include file="include/mmd5.asp"-->
<%
product_info_name=request("product_info_name")
product_info_priceS=request("product_info_priceS")
Url=request("url")
flag=request("flag")   '1ʱ�����ͷ���,�ޱ�ʾ����֧������ҳ����,�������ͷ���

sql="select base_NetPay_AlipayEmail,base_NetPay_AlipaySafeCode,base_NetPay_AlipayPartnerID from root_NetPay where base_NetPay_id=1"
set rs=conn.execute (sql)
base_NetPay_AlipayEmail   =rs(0)
base_NetPay_AlipaySafeCode=rs(1)
base_NetPay_AlipayPartnerID=rs(2)
rs.close
set rs=nothing
      
if flag=1 then
    deliverprice1=5   'ƽ�ʷ����趨(ƥ��֧������)
    deliverprice2=20  '������ݷ����趨(ƥ��֧������)
else
    deliverprice1=""   'ƽ�ʷ����趨(ƥ��֧������)
    deliverprice2=""   '������ݷ����趨(ƥ��֧������)
end if

v_mid=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipayEmail))),chr(13)&chr(10),"<br>")
key=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipaySafeCode))),chr(13)&chr(10),"<br>")
partnerID=replace(strAnsi2Unicode(Base64decode(strUnicode2Ansi(base_NetPay_AlipayPartnerID))),chr(13)&chr(10),"<br>")

s1           =        "0001"		  '������
s2           =        server.urlEncode(product_info_name)	  '��Ʒ����
s3           =        ""
s4           =        product_info_priceS
s22           =       url
s21          =        server.urlEncode(product_info_name)	  '��Ʒ����
s5           =        "1"		  '֧������
s6           =        ""		  '��������
s7           =        ""		  '������ʽ
s8           =        deliverprice1
s9           =        deliverprice2
s10          =        "true"        	  'ֻ��
s11          =        content		  '�������
s12          =        ""		  '���Email
s13          =        ucase(order_info_realname)	  '�������
s14          =        order_info_address	  '��ҵ�ַ
s15          =        order_info_zip		  '����ʱ�
s17          =        order_info_mobile		  '����ֻ�
sellerEmail  =        v_mid		  '����EMAIL
s18    	     =        partnerID  'partner
key          =        key
str2CreateAc = "cmd" & s1 & "subject" & s2
str2CreateAc = str2CreateAc & "body" & s21
str2CreateAc = str2CreateAc & "order_no" & s3
str2CreateAc = str2CreateAc & "price" & s4
str2CreateAc = str2CreateAc & "url" & s22
str2CreateAc = str2CreateAc & "type" & s5
str2CreateAc = str2CreateAc & "number" & s6
str2CreateAc = str2CreateAc & "transport" & s7
str2CreateAc = str2CreateAc & "ordinary_fee" & s8
str2CreateAc = str2CreateAc & "express_fee" & s9
str2CreateAc = str2CreateAc & "readonly" & s10
str2CreateAc = str2CreateAc & "buyer_msg" & s11
str2CreateAc = str2CreateAc & "seller" & sellerEmail
str2CreateAc = str2CreateAc & "buyer" & s12
str2CreateAc = str2CreateAc & "buyer_name" & s13
str2CreateAc = str2CreateAc & "buyer_address" & s14
str2CreateAc = str2CreateAc & "buyer_zipcode" & s15
str2CreateAc = str2CreateAc & "buyer_tel" & s16
str2CreateAc = str2CreateAc & "buyer_mobile" & s17
str2CreateAc = str2CreateAc & "partner" & s18
str2CreateAc = str2CreateAc & key
	
ac=MD5(str2CreateAc)

response.write ("<form method=post name=post action=https://www.alipay.com/payto:"&v_mid&">")
response.write ("<input type=hidden name=cmd value="&s1&">")
response.write ("<input type=hidden name=subject value="&s2&">")
response.write ("<input type=hidden name=body value="&s21&">")
response.write ("<input type=hidden name=order_no value="&s3&">")
response.write ("<input type=hidden name=price value="&s4&">")
response.write ("<input type=hidden name=url value="&s22&">")
response.write ("<input type=hidden name=type value="&s5&">")
response.write ("<input type=hidden name=number value="&s6&">")
response.write ("<input type=hidden name=transport value="&s7&">")
response.write ("<input type=hidden name=ordinary_fee value="&s8&">")
response.write ("<input type=hidden name=express_fee value="&s9&">")
response.write ("<input type=hidden name=readonly value="&s10&">")
response.write ("<input type=hidden name=buyer_msg value="&s11&">")
response.write ("<input type=hidden name=seller value="&sellerEmail&">")
response.write ("<input type=hidden name=buyer value="&s12&">")
response.write ("<input type=hidden name=buyer_name value="&s13&">")
response.write ("<input type=hidden name=buyer_address value="&s14&">")
response.write ("<input type=hidden name=buyer_zipcode value="&s15&">")
response.write ("<input type=hidden name=buyer_tel value="&s16&">")
response.write ("<input type=hidden name=buyer_mobile value="&s17&">")
response.write ("<input type=hidden name=partner value="&s18&">")
response.write ("<input type=hidden name=ac value="&ac&">")
response.write ("</form>")
response.write "<SCRIPT LANGUAGE=""JavaScript"">"
response.write "document.post.submit();"
response.write "</SCRIPT>"
%>
</center>