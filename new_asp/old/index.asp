<!--#include file="alipay/alipay_payto.asp"-->
<%
dim service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,seller_email,notify_url,return_url
dim t1,t4,t5,key
dim AlipayObj,itemUrl
if not isnumeric(Request.QueryString("prices")) then

response.write("����,������׼ȷ���")
response.end
end if
	t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'֧���ӿ�
	t4				=	"images/alipay_bwrx.gif"		'֧������ťͼƬ
	t5				=	"�Ƽ�ʹ��֧��������"						'��ť��ͣ˵��
	
	service         =   "trade_create_by_buyer"
	agent           =   "citemn@yahoo.cn" '��������id
	partner			=	trim(replace(request("id"),"'",""))		'partner�������ID(������)
	sign_type       =   "MD5"
	subject			=	replace(trim(request("subject")),"'","")		'��Ʒ����
	body			=	replace(trim(request("body")),"'","")	'body			��Ʒ����
	'out_trade_no    =   Replace(Now(),"-","")           '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	out_trade_no    =   trim(replace(request("order_id"),"'","")) '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	price		    =	Request.QueryString("Memony")				'price��Ʒ����			0.01��50000.00
    discount        =   "0"               '��Ʒ�ۿ�
    show_url        =   ""        '��Ʒչʾ��ַ������ֱ��д��վ��ҳ��ַ��
    quantity        =   trim(request("product_count"))               '��Ʒ����
    payment_type    =   "1"                '֧�����ͣ���1������Ʒ����
    logistics_type  =   "EXPRESS"          '�������ࣨ��ݣ�
    logistics_fee   =   trim(request("yinfei"))                '��������
    logistics_payment  =   "SELLER_PAY"    '�������óе�(���Ҹ�)
	logistics_type_1  =   "EMS"
    logistics_fee_1   =   trim(request("yinfei"))
    logistics_payment_1  =   "BUYER_PAY"   '�������óе�(��Ҹ�)
    seller_email    =  Trim(Request("seller_email"))   '(������)
    key             =    replace(trim(request("Key")),"'","")  '(������)
    notify_url=  Trim(Request("return_url"))   '������֪ͨurl����ʹ�ã��벻Ҫע�ͻ���ɾ���˲��������ô��ݸ�֧����ϵͳ��Alipay_Notify.asp�ļ�����·���� 
    return_url=  Trim(Request("return_url"))   '������֪ͨurl����ʹ�ã��벻Ҫע�ͻ���ɾ���˲��������ô��ݸ�֧����ϵͳ��Alipay_Notify.asp�ļ�����·���� 

	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(t1,t4,t5,service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,seller_email,notify_url,return_url,key)
	
	dim zifubao_url
	zifubao_url=""
	zifubao_url=mid(itemUrl,InStr(itemUrl,"https://"),InStr(itemUrl,"target='_blank'")-12)
	response.Redirect(zifubao_url)
%>
