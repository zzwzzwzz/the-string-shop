<%
'====================   支 付 宝   ====================
sub pay_AliPay()
	dim service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,receive_name,seller_email,notify_url
	dim t1,t4,t5,key
	dim AlipayObj,itemUrl
	dim alipay_name
    alipay_name=order_info_no&"号订单"

	t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'支付接口
	t4				=	"images/alipay_bwrx.gif"		'支付宝按钮图片
	t5				=	"推荐使用支付宝付款"						'按钮悬停说明
	
	service         =   "trade_create_by_buyer"
	agent           =   "" '合作厂商id
	partner			=	PartnerID		'partner合作伙伴ID(必须填)
	sign_type       =   "MD5"
    subject         =   alipay_name	  '商品名称
	body			=	ProdNames&"配送费用已包含在内"		'body			商品描述
	out_trade_no    =   order_info_no           '客户网站订单号，（现取系统时间，可改成网站自己的变量）
	price		    =	order_info_allcost				'price商品单价			0.01～50000.00
    discount        =   "0"               '商品折扣
    show_url        =   ""        '商品展示地址（可以直接写网站首页网址）
    quantity        =   "1"               '商品数量
    payment_type    =   "1"                '支付类型，（1代表商品购买）
    logistics_type  =   "EXPRESS"          '物流种类（快递）
    logistics_fee   =   "0"                '物流费用
    logistics_payment  =   "SELLER_PAY"    '物流费用承担(卖家付)
	logistics_type_1  =   ""
    logistics_fee_1   =   "0"
    logistics_payment_1  =   "BUYER_PAY"   '物流费用承担(买家付)
    seller_email    =    v_mid   '(必须填)
    key             =    key1  '(必须填)
    
    notify_url=  "alipay/Alipay_Notify.asp"   '服务器通知url（不使用，请不要注释或者删除此参数，不用传递给支付宝系统，Alipay_Notify.asp文件所在路经） 

	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(t1,t4,t5,service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,seller_email,notify_url,key)	
	response.write itemUrl
end sub

'====================    PayPal    ====================
sub pay_paypal()
    response.write ("<form action=https://www.paypal.com/cgi-bin/webscr method=post name=post>")
    response.write ("<input type=hidden name=cmd value=_xclick>")
    response.write ("<input type=hidden name=business value="&v_mid&">")
    response.write ("<input type=hidden name=return value=../PayBackPayPal.asp>")
    response.write ("<input type=hidden name=item_name value=Order NO. is "&Order_info_No&">")
    response.write ("<input type=hidden name=item_number value="&Order_info_No&">")
    response.write ("<input type=hidden name=amount value="&order_info_AllCost&">")
    response.write ("</form>")
    response.write "<SCRIPT LANGUAGE=""JavaScript"">"
    response.write "document.post.submit();"
    response.write "</SCRIPT>"
end sub

%>
 
