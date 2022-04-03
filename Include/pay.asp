<%
'====================   Ipay   ====================
sub pay_ipay()
    v_oid      =Order_info_No        '订单号
    v_amount   =order_info_AllCost   '总价
    v_email    =""                   '付款人电子邮件
    v_mobile   =order_info_mobile    '付款人联系电话
    v_md5      =md5(v_mid & v_oid & v_amount & v_email & v_mobile & key)  '组合md5加密
    v_comment_1="购物款项" 
    v_comment  =server.urlencode(v_comment_1)
    v_url      ="../PayBack_IPay.asp"  '支付结果后返回页
    response.write ("<form method=post name=post action=""http://www.ipay.cn/4.0/bank.shtml"">")
    response.write ("<input type=hidden name=v_mid value="&v_mid&">")
    response.write ("<input type=hidden name=v_oid value="&v_oid&">")
    response.write ("<input type=hidden name=v_amount value="&v_amount&">")
    response.write ("<input type=hidden name=v_md5 value="&v_md5&">")
    response.write ("<input type=hidden name=v_email value="&v_email&">")
    response.write ("<input type=hidden name=v_mobile value="&v_mobile&">")
    response.write ("<input type=hidden name=v_comment value="&v_comment&">")
    response.write ("<input type=hidden name=v_url  value="&v_url&">") 
    response.write ("</form>")
    response.write "<SCRIPT LANGUAGE=""JavaScript"">"
    response.write "document.post.submit();"
    response.write "</SCRIPT>"
end sub


'====================   网银在线   ====================
sub pay_ChinaBank()
    v_oid   =order_info_no
    v_amount=order_info_AllCost
    v_email =order_info_email
    v_mobile=order_info_tel

    v_md5=md5(v_mid & v_oid & v_amount & v_email & v_mobile & key)
    v_url="../PayBack_ChinaBank.asp"
    v_moneytype = "0"
    style="0"
    remark1=""
    remark2=""
    text = v_amount&v_moneytype&v_oid&v_mid&v_url&key
    v_md5info=Ucase(trim(md5(text)))	
    response.write ("<form method=post action=https://pay.chinabank.com.cn/select_bank name=E_FORM target=new>")
    response.write ("<input type=hidden name=v_md5info size=100  value="&v_md5info&">")
    response.write ("<input type=hidden name=v_mid value="&v_mid&">")
    response.write ("<input type=hidden name=v_oid value="&v_oid&">")
    response.write ("<input type=hidden name=v_amount value="&v_amount&">")
    response.write ("<input type=hidden name=v_moneytype  value="&v_moneytype&">")
    response.write ("<input type=hidden name=v_url value="&v_url&">")
    response.write ("<input type=hidden name=style value="&style&">")
    response.write ("<input type=hidden name=remark1 value="&remark1&">")
    response.write ("<input type=hidden name=remark2 value="&remark2&">")
    response.write ("</form>")
    response.write "<SCRIPT LANGUAGE=""JavaScript"">"
    response.write "document.E_FORM.submit();"
    response.write "</SCRIPT>"
end sub

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



'====================     NPS     ====================
sub pay_NPS()
	m_id		=	v_mid
	modate		=	date()
	m_orderid	=	Order_info_No
	m_oamount	=	order_info_AllCost
	m_ocurrency	=	"1"
	m_url		=	"../PayBack_NPS.asp"
	m_language	=	"1"
	s_name		=	order_info_RealName
	s_addr		=	order_info_address
	s_postcode	=	order_info_zip
	s_tel		=	order_info_tel			
	s_eml		=	""
	r_name		=	order_info_RealName
	r_addr		=	order_info_address
	r_postcode	=	order_info_zip
	r_tel		=	order_info_tel
	r_eml		=	""
	m_ocomment	=	"商品购买"	
	m_status	=	"0"
	key		    =	key
	
	OrderMessage =m_id&m_orderid&m_oamount&m_ocurrency&m_url&m_language&s_postcode&s_tel&s_eml&r_postcode&r_tel&r_eml&modate&key
	
	digest = Ucase(trim(md5(OrderMessage)))
	
	response.write ("<form method=get name=post action=https://payment.nps.cn/VirReceiveMerchantAction.do>")
	response.write ("<input Type=hidden Name=M_ID value="&m_id&">")
	response.write ("<input Type=hidden Name=MOrderID value="&m_orderid&">")
	response.write ("<input Type=hidden Name=MOAmount value="&m_oamount&">")
	response.write ("<input Type=hidden Name=MOCurrency value="&m_ocurrency&">")
	response.write ("<input Type=hidden Name=M_URL value="&m_url&">")
	response.write ("<input Type=hidden Name=M_Language value="&m_language&">")
	response.write ("<input Type=hidden Name=S_Name value="&s_name&">")
	response.write ("<input Type=hidden Name=S_Address value="&s_addr&">")
	response.write ("<input Type=hidden Name=S_PostCode value="&s_postcode&">")
	response.write ("<input Type=hidden Name=S_Telephone value="&s_tel&">")
	response.write ("<input Type=hidden Name=S_Email value="&s_eml&">")
	response.write ("<input Type=hidden Name=R_Name value="&r_name&">")
	response.write ("<input Type=hidden Name=R_Address value="&r_addr&">")
	response.write ("<input Type=hidden Name=R_PostCode value="&r_postcode&">")
	response.write ("<input Type=hidden Name=R_Telephone value="&r_tel&">")
	response.write ("<input Type=hidden Name=R_Email value="&r_eml&">")
	response.write ("<input Type=hidden Name=MOComment value="&m_ocomment&">")
	response.write ("<input Type=hidden Name=MODate value="&modate&">")
	response.write ("<input Type=hidden Name=State value="&m_status&">")
	response.write ("<input Type=hidden Name=digestinfo value="&digest&">")
	response.write ("</form>")
	
	response.write "<SCRIPT LANGUAGE=""JavaScript"">"
	response.write "document.post.submit();"
	response.write "</SCRIPT>"
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
 
