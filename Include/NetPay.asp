<%
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
    
    notify_url=  ""   '������֪ͨurl����ʹ�ã��벻Ҫע�ͻ���ɾ���˲��������ô��ݸ�֧����ϵͳ��Alipay_Notify.asp�ļ�����·���� 
    
    itemUrl=creatAlipayURL(t1,t4,t5,service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,seller_email,notify_url,key)
end sub
%>
 
