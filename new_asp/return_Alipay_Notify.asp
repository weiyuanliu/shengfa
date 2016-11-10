<%
	'功能：付完款后跳转的页面
	'版本：2.0
	'日期：2008-1-5
	'作者：支付宝公司销售部技术支持团队
	'联系：0571-26888888
	'版权：支付宝公司
%>

<!--#include file="alipayto/Alipay_md5.asp"-->
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->

<%
	  partner			="2088402551356533"	 '填写对应支付宝账户的合作者身份ID
	  key			    ="dinrqtpwtcai6wzv4iy8qby016hb67uo"	 '填写对应支付宝帐户的安全校验码

	out_trade_no	= DelStr(Request("out_trade_no"))  '获取定单号
    total_fee		= DelStr(Request("total_fee"))     '获取支付的总价格
	'如需获取其它参数，可填写 参数 =DelStr(Request.Form("获取参数名"))

'**********************判断消息是不是支付宝发出********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    '暂时注释下面两行2013-8-29
    'Retrieval.send()
    'ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*******************************************************************

'*******获取支付宝GET过来通知消息,判断消息是不是被修改过************
For Each varItem in Request.QueryString
	mystr=varItem&"="&Request(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
	mystr=Left(mystr,Len(mystr)-1)
End If 
mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'对参数排序
For i = Count TO 0 Step -1
	minmax = mystr( 0 )
	minmaxSlot = 0
	For j = 1 To i
		mark = (mystr( j ) > minmax)
		If mark Then 
			minmax = mystr( j )
			minmaxSlot = j
		End If 
	Next    
	If minmaxSlot <> i Then 
		temp = mystr( minmaxSlot )
		mystr( minmaxSlot ) = mystr( i )
		mystr( i ) = temp
	End If
Next
'构造md5摘要字符串
For j = 0 To Count Step 1
	value = SPLIT(mystr( j ), "=")
	If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
		If j=Count Then
			md5str= md5str&mystr( j )
		Else 
			md5str= md5str&mystr( j )&"&"
		End If 
	End If 
Next
md5str=md5str&key
mysign=md5(md5str)
'********************************************************

If mysign=Request("sign") Then 	
 	call SuccesOrder(out_trade_no)
	call AboutShow(54,out_trade_no,total_fee)
	'call sendSms(4,rs("Linkman"),rs("Tel")) 
Else
	response.write "也许您支付成功，我们的网站却没有记录您的数据，请联系客服确认!"          '这里可以指定你需要显示的内容
End If 


Function DelStr(Str)
	If IsNull(Str) Or IsEmpty(Str) Then
		Str	= ""
	End If
	DelStr	= Replace(Str,";","")
	DelStr	= Replace(DelStr,"'","")
	DelStr	= Replace(DelStr,"&","")
	DelStr	= Replace(DelStr," ","")
	DelStr	= Replace(DelStr,"　","")
	DelStr	= Replace(DelStr,"%20","")
	DelStr	= Replace(DelStr,"--","")
	DelStr	= Replace(DelStr,"==","")
	DelStr	= Replace(DelStr,"<","")
	DelStr	= Replace(DelStr,">","")
	DelStr	= Replace(DelStr,"%","")
End Function

Function SuccesOrder(OrderId)
 Dim rs,sql
 set rs=server.CreateObject("Adodb.recordset")
 sql="Select * from NwebCn_Order where ProductNo = '"&OrderId&"'"
 rs.open sql,conn,1,3
 if not rs.eof then
  rs("State")="货款已付"
 rs.update
 end if
 rs.close
 set rs=nothing
End Function

Function AboutShow(Id,out_trade_no,total_fee)
 Dim rs,sql,Text
 set rs=server.CreateObject("Adodb.recordset")
 sql="Select * from NwebCn_About where Id="&Id&""
 rs.open sql,conn,1,3
 if not rs.eof then
  Text=rs("Content")
  Text=replace(Text,"{订单编号}",out_trade_no)
  Text=replace(Text,"{支付金额}",total_fee)
  response.Write Text
  rs("ClickNumber")=rs("ClickNumber")+1
  rs.update
 end if
 rs.close
 set rs=nothing
End Function

%>