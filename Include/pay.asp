<%
'====================   ֧ �� ��   ====================
sub pay_AliPay()
	dim service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,receive_name,seller_email,notify_url
	dim t1,t4,t5,key
	dim AlipayObj,itemUrl
	dim alipay_name
    alipay_name=order_info_no&"�Ŷ���"

	t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'֧���ӿ�
	t4				=	"images/alipay_bwrx.gif"		'֧������ťͼƬ
	t5				=	"�Ƽ�ʹ��֧��������"						'��ť��ͣ˵��
	
	service         =   "trade_create_by_buyer"
	agent           =   "" '��������id
	partner			=	PartnerID		'partner�������ID(������)
	sign_type       =   "MD5"
    subject         =   alipay_name	  '��Ʒ����
	body			=	ProdNames&"���ͷ����Ѱ�������"		'body			��Ʒ����
	out_trade_no    =   order_info_no           '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	price		    =	order_info_allcost				'price��Ʒ����			0.01��50000.00
    discount        =   "0"               '��Ʒ�ۿ�
    show_url        =   ""        '��Ʒչʾ��ַ������ֱ��д��վ��ҳ��ַ��
    quantity        =   "1"               '��Ʒ����
    payment_type    =   "1"                '֧�����ͣ���1������Ʒ����
    logistics_type  =   "EXPRESS"          '�������ࣨ��ݣ�
    logistics_fee   =   "0"                '��������
    logistics_payment  =   "SELLER_PAY"    '�������óе�(���Ҹ�)
	logistics_type_1  =   ""
    logistics_fee_1   =   "0"
    logistics_payment_1  =   "BUYER_PAY"   '�������óе�(��Ҹ�)
    seller_email    =    v_mid   '(������)
    key             =    key1  '(������)
    
    notify_url=  "alipay/Alipay_Notify.asp"   '������֪ͨurl����ʹ�ã��벻Ҫע�ͻ���ɾ���˲��������ô��ݸ�֧����ϵͳ��Alipay_Notify.asp�ļ�����·���� 

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
 
