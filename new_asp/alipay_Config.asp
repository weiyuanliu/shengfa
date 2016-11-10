<%
	'功能：设置商户有关信息及返回路径
	'版本：2.0
	'日期：2009-1-5
	'作者：支付宝公司销售部技术支持团队
	'联系：0571-26888888
	'版权：支付宝公司

'/*****************************************
'登陆 签约支付宝账号（www.alipay.com） 后-------->商家服务,可以看到合作者身份（partnerID）和交易安全校验码（key）
'******************************************
	
	  
	  
      show_url          ="http://www.beloj.com"  '商户网站的网址,例如:www.alipay.com
	  seller_email		="beloj@aliyun.com"	 '请填写支付宝签约帐户
	  partner			="2088402551356533"	 '填写对应支付宝账户的合作者身份ID
	  key			    ="dinrqtpwtcai6wzv4iy8qby016hb67uo"	 '填写对应支付宝帐户的安全校验码
	  return_url		="http://www.beloj.com/new_asp/return_Alipay_Notify.asp"
      notify_url		="http://www.beloj.com/new_asp/Alipay_Notify.asp"	      
	  '交易过程中服务器通知的页面，例如http://www.alipay.com/alipay/Alipay_Notify.asp  注意是文件的绝对路径。
	  'return_url		="http://citemn.com/alipay1/return_Alipay_Notify.asp"  
	  '付完款后跳转的页面，例如http://www.alipay.com/alipay/return_Alipay_Notify.asp   注意是文件的绝对路径。
	  '如果使用了Alipay_Notify.asp或者return_Alipay_Notify.asp，请在这两个文件中添加相应的合作身份者ID和安全校验码

%>