<!--#include file="alipay/alipay_payto.asp"-->
<%
dim service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,seller_email,notify_url,return_url
dim t1,t4,t5,key
dim AlipayObj,itemUrl
if not isnumeric(Request.QueryString("prices")) then

response.write("错误,请输入准确金额")
response.end
end if
	t1				=	"https://www.alipay.com/cooperate/gateway.do?"	'支付接口
	t4				=	"images/alipay_bwrx.gif"		'支付宝按钮图片
	t5				=	"推荐使用支付宝付款"						'按钮悬停说明
	
	service         =   "trade_create_by_buyer"
	agent           =   "citemn@yahoo.cn" '合作厂商id
	partner			=	trim(replace(request("id"),"'",""))		'partner合作伙伴ID(必须填)
	sign_type       =   "MD5"
	subject			=	replace(trim(request("subject")),"'","")		'商品名称
	body			=	replace(trim(request("body")),"'","")	'body			商品描述
	'out_trade_no    =   Replace(Now(),"-","")           '客户网站订单号，（现取系统时间，可改成网站自己的变量）
	out_trade_no    =   trim(replace(request("order_id"),"'","")) '客户网站订单号，（现取系统时间，可改成网站自己的变量）
	price		    =	Request.QueryString("Memony")				'price商品单价			0.01～50000.00
    discount        =   "0"               '商品折扣
    show_url        =   ""        '商品展示地址（可以直接写网站首页网址）
    quantity        =   trim(request("product_count"))               '商品数量
    payment_type    =   "1"                '支付类型，（1代表商品购买）
    logistics_type  =   "EXPRESS"          '物流种类（快递）
    logistics_fee   =   trim(request("yinfei"))                '物流费用
    logistics_payment  =   "SELLER_PAY"    '物流费用承担(卖家付)
	logistics_type_1  =   "EMS"
    logistics_fee_1   =   trim(request("yinfei"))
    logistics_payment_1  =   "BUYER_PAY"   '物流费用承担(买家付)
    seller_email    =  Trim(Request("seller_email"))   '(必须填)
    key             =    replace(trim(request("Key")),"'","")  '(必须填)
    notify_url=  Trim(Request("return_url"))   '服务器通知url（不使用，请不要注释或者删除此参数，不用传递给支付宝系统，Alipay_Notify.asp文件所在路经） 
    return_url=  Trim(Request("return_url"))   '服务器通知url（不使用，请不要注释或者删除此参数，不用传递给支付宝系统，Alipay_Notify.asp文件所在路经） 

	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(t1,t4,t5,service,agent,partner,sign_type,subject,body,out_trade_no,price,discount,show_url,quantity,payment_type,logistics_type,logistics_fee,logistics_payment,logistics_type_1,logistics_fee_1,logistics_payment_1,seller_email,notify_url,return_url,key)
	
	dim zifubao_url
	zifubao_url=""
	zifubao_url=mid(itemUrl,InStr(itemUrl,"https://"),InStr(itemUrl,"target='_blank'")-12)
	response.Redirect(zifubao_url)
%>
