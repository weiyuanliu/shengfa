<%
	'版本：2.0
	'日期：2009-01-05
	'作者：支付宝公司销售部技术支持团队
	'联系：0571-26888888
	'版权：支付宝公司
%>

<!--#include file="alipayto/alipay_payto.asp"-->
<%
   shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '客户网站订单号，（现取系统时间，可改成网站自己的变量）
	
	subject			=	Trim(Request("subject"))	'商品名称，如果客户走购物车流程可以设为  "订单号："&request("客户网站订单")
	body			=	Trim(Request("body"))		'商品描述
	out_trade_no    =   Trim(Request("order_id"))         '按时间获取的订单号
	price		    =	Trim(Request("Memony"))			'price商品单价	0.01～50000.00 , 注：不要出现3,000.00，价格不支持","号
    quantity        =   "1"             '商品数量,如果走购物车默认为1
	discount        =   "0"             '商品折扣
    seller_email    =    seller_email   '卖家的支付宝帐号,c2c客户，可以更改此参数。
	paymethod       =   "directPay"      '赋值:bankPay(网银);cartoon(卡通); directPay(余额)
	defaultbank     =   "directPay"     ' 网银默认的银行
	Set AlipayObj	=   New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email,paymethod)
	response.Redirect(itemUrl)
%>