<%
	'�汾��2.0
	'���ڣ�2009-01-05
	'���ߣ�֧������˾���۲�����֧���Ŷ�
	'��ϵ��0571-26888888
	'��Ȩ��֧������˾
%>

<!--#include file="alipayto/alipay_payto.asp"-->
<%
   shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	
	subject			=	Trim(Request("subject"))	'��Ʒ���ƣ�����ͻ��߹��ﳵ���̿�����Ϊ  "�����ţ�"&request("�ͻ���վ����")
	body			=	Trim(Request("body"))		'��Ʒ����
	out_trade_no    =   Trim(Request("order_id"))         '��ʱ���ȡ�Ķ�����
	price		    =	Trim(Request("Memony"))			'price��Ʒ����	0.01��50000.00 , ע����Ҫ����3,000.00���۸�֧��","��
    quantity        =   "1"             '��Ʒ����,����߹��ﳵĬ��Ϊ1
	discount        =   "0"             '��Ʒ�ۿ�
    seller_email    =    seller_email   '���ҵ�֧�����ʺ�,c2c�ͻ������Ը��Ĵ˲�����
	paymethod       =   "directPay"      '��ֵ:bankPay(����);cartoon(��ͨ); directPay(���)
	defaultbank     =   "directPay"     ' ����Ĭ�ϵ�����
	Set AlipayObj	=   New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email,paymethod)
	response.Redirect(itemUrl)
%>