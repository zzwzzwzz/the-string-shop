<!--#include file="base64.asp"-->
<!--#include file="mmd5.asp"-->
<%
call alipay()
sub alipay()
	INTERFACE_URL = "https://www.alipay.com/payto:"  '支付接口
  	sellerEmail  = "520@163.com"  '商户支付宝账户（改成你自己的）
  	keyCode   = "8tucxmhlxfx3n"  '安全校验码（改成你自己的）
  	cmd    = "0001"  '命令码
  	subject   = "商品名称"   '商品名称
  	body   = "无"  '商品描述
  	prices   = "0.01"  '商品单价
  	order_no  = "订单号"  '商户订单号
  	types  = 1    'type支付类型  1：商品购买2：服务购买3：网络拍卖4：捐赠
  	number   = "1" '购买数量
  	transport  = 3    '发货方式   1：平邮2：快递3：虚拟物品
  	ordinary_fee = 0    '平邮运费
  	express_fee  = 0    '快递运费
  	buyer   = "njjmail@126.com"   'buyer   买家Email
  	buyer_name  = "姓名"  '买家姓名
  	buyer_address = "上海"  '买家地址
  	buyer_zipcode = "311500"  '买家邮编
  	buyer_tel  = "05718927377"  '买家电话
  	partner   = "208804871084"  '合作伙伴

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