<!--#include file="base64.asp"-->
<!--#include file="mmd5.asp"-->
<%
call alipay()
sub alipay()
	INTERFACE_URL = "https://www.alipay.com/payto:"  '֧���ӿ�
  	sellerEmail  = "520@163.com"  '�̻�֧�����˻����ĳ����Լ��ģ�
  	keyCode   = "8tucxmhlxfx3n"  '��ȫУ���루�ĳ����Լ��ģ�
  	cmd    = "0001"  '������
  	subject   = "��Ʒ����"   '��Ʒ����
  	body   = "��"  '��Ʒ����
  	prices   = "0.01"  '��Ʒ����
  	order_no  = "������"  '�̻�������
  	types  = 1    'type֧������  1����Ʒ����2��������3����������4������
  	number   = "1" '��������
  	transport  = 3    '������ʽ   1��ƽ��2�����3��������Ʒ
  	ordinary_fee = 0    'ƽ���˷�
  	express_fee  = 0    '����˷�
  	buyer   = "njjmail@126.com"   'buyer   ���Email
  	buyer_name  = "����"  '�������
  	buyer_address = "�Ϻ�"  '��ҵ�ַ
  	buyer_zipcode = "311500"  '����ʱ�
  	buyer_tel  = "05718927377"  '��ҵ绰
  	partner   = "208804871084"  '�������

  	str2CreateAc = "cmd" & cmd & "subject" & subject
  	str2CreateAc = str2CreateAc & "body" & body
  	str2CreateAc = str2CreateAc & "order_no" & order_no
  	str2CreateAc = str2CreateAc & "price" & prices
  	str2CreateAc = str2CreateAc & "type" & types
  	str2CreateAc = str2CreateAc & "number" & number
  	str2CreateAc = str2CreateAc & "transport" & transport
  	str2CreateAc = str2CreateAc & "ordinary_fee" & ordinary_fee
  	str2CreateAc = str2CreateAc & "express_fee" & express_fee
  	str2CreateAc = str2CreateAc & "seller" & sellerEmail
	'  str2CreateAc = str2CreateAc & "buyer" & buyer  
  	str2CreateAc = str2CreateAc & "buyer_name" & buyer_name
  	str2CreateAc = str2CreateAc & "buyer_address" & buyer_address
  	str2CreateAc = str2CreateAc & "buyer_zipcode" & buyer_zipcode
  	str2CreateAc = str2CreateAc & "buyer_tel" & buyer_tel
  	str2CreateAc = str2CreateAc & "partner" & partner
  	str2CreateAc = str2CreateAc & keyCode
  	
  	ac=MD5(str2CreateAc)

  	itemURL   =  INTERFACE_URL & sellerEmail & "?cmd=" & cmd
  	itemURL   =  itemURL & "&subject=" & Server.urlEncode(subject)
  	itemURL   =  itemURL & "&body=" & Server.urlEncode(body)
  	itemURL   =  itemURL & "&order_no=" & order_no
  	itemURL   =  itemURL & "&price=" & prices
  	itemURL   =  itemURL & "&type=" & types
  	itemURL   =  itemURL & "&number=" & number
  	itemURL   =  itemURL & "&transport=" & transport
  	itemURL   =  itemURL & "&ordinary_fee=" & ordinary_fee
 	itemURL   =  itemURL & "&express_fee=" & express_fee
  	itemURL   =  itemURL & "&seller=" & Server.urlEncode(sellerEmail)
	'  itemURL   =  itemURL & "&buyer=" & Server.HTMLEncode(buyer)  
  	itemURL   =  itemURL & "&buyer_name=" & Server.urlEncode(buyer_name)
  	itemURL   =  itemURL & "&buyer_address=" & Server.urlEncode(buyer_address)
  	itemURL   =  itemURL & "&buyer_zipcode=" & Server.urlEncode(buyer_zipcode)
  	itemURL   =  itemURL & "&buyer_tel=" & buyer_tel
  	itemURL   =  itemURL & "&partner=" & partner
  	itemURL   =  itemURL & "&ac=" & ac

	'response.write(itemURL&str2CreateAc)
	'response.End()
	response.Redirect(itemURL)
end sub
%>