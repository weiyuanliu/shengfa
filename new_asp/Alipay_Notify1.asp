<%
	'功能：付款过程中服务器通知页面
	'版本：2.0
	'日期：2008-1-5
	'作者：支付宝公司销售部技术支持团队
	'联系：0571-26888888
	'版权：支付宝公司
%>

<!--#include file="alipayto/Alipay_md5.asp"-->
<%
    key="7kyhcjza17shaxiutofguau6kjryinti"         '支付宝安全教研码
    partner="2088102160488222"     '支付宝合作id 
 
	out_trade_no	=DelStr(Request.Form("out_trade_no"))      '获取定单号
    total_fee		=DelStr(Request.Form("total_fee"))         '获取支付的总价格
	'如需获取其它参数，可填写 参数 =DelStr(Request.Form("获取参数名"))
	
'*******************判断消息是不是支付宝发出***********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request.Form("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*******************************************************************

'*******************获取支付宝POST过来通知消息**********************
For Each varItem in Request.Form
	mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
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

'*************************交易状态返回处理*************************
If mysign=request.Form("sign") And ResponseTxt="true" Then 	
	If request.Form("trade_status") = "TRADE_FINISHED" Then 
' 如果您申请了支付宝的购物卷功能，请在返回的信息里面不要做金额的判断，否则会出现校验通不过，出现调单。如果您需要获取买家所使用购物卷的金额,
' 请获取返回信息的这个字段discount的值，取绝对值，就是买家付款优惠的金额。即 原订单的总金额=买家付款返回的金额total_fee +|discount|.
		'在此处添加：付款成功,更新数据库语句  
	 	if AliplaySuccess(out_trade_no) then '支付成功的处理过程
			ShowMsg "提示信息","恭喜您，在线支付成功！"
		else
			ShowErrorMsg "提示信息","支付失败，请返回！"	
		end if
	Else '支付失败执行的程序
		ShowErrorMsg "提示信息","支付失败，请返回！"			
	End If
	'Response.Write returnTxt输出交易返回的太态
Else
	ShowErrorMsg "提示信息","数据出错，支付失败！请返回！"	
	'response.write "fail" '获取数据出错的显示信息
End If 
'*******************************************************************

'如果服务器支持文本写入日志，那么可以打开下边注释部分，方便测试。

 '写文本，方便测试（看网站需求，也可以改成存入数据库）
'TOEXCELLR=TOEXCELLR&md5str&"MD5结果:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
'set fs= createobject("scripting.filesystemobject") 
'set ts=fs.createtextfile(server.MapPath("alipayto/Notify_DATA/"&replace(now(),":","")&".txt"),true)

' ts.writeline(TOEXCELLR)
 'ts.close
' set ts=Nothing
' set fs=Nothing

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


'////自定义函数
Function AliplaySuccess(OrderID) '支付成后的处理程序
	if OrderID <> "" then
		Dim conn,rs,sql
		CreateConn Conn '建立链接对像
		CreateRs rs '建立记录集对像
		
		Sql="Select State from NwebCn_Order where ProductNo='"&OrderID&"' "
		rs.open sql,conn,1,3
		if rs("State")="" or rs("state")=null Then
		if rs.eof and rs.bof then
			AliplaySuccess=False
		else
			rs("State")="货款已付"
			AliplaySuccess=True
			Rs.update()
		end if
		end if
		CloseObject rs
		CloseObject Conn
	end if
End Function

'创建 Conn对像
Sub CreateConn(ByRef Conn)
	Dim ConnStr
	On error resume next
	Set Conn=Server.CreateObject("Adodb.Connection")
	ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("../Database/NwebCn_Site.asp")
	Conn.open ConnStr
	if err then
	   err.clear
	   Set Conn = Nothing
	   Response.Write "系统错误：数据库连接出错，请检查'系统管理>>站点常量设置',或者/Include/Const.asp文件!"
	   Response.End
	end if
End Sub

'创建记录集对像
Sub CreateRs(ByRef Object)
	Set Object=server.CreateObject("Adodb.Recordset")
End Sub

Sub CloseObject(ByRef Object)
	Object.Close()
	Set Object=Nothing
End Sub
%>
<!--提示信息-->
<%Sub ShowMsg(Title,Text)%>
<link rel="stylesheet" href="dxdiag.css" />
<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td style="border:#C1C1C1 1px solid;">
<table width="227" border="0" cellspacing="0">
  <tr>
    <td height="26" bgcolor="#2A7FFF" style="color:#FFFFFF; padding-left:10px; font-weight:bold;" class="DxdiagTitel">提示信息：</td>
  </tr>
  <tr>
    <td height="51" bgcolor="#EFEFEF" style="padding-left:10px; padding-right:10px;" class="DxdiagText"><%=Text%></td>
  </tr>
  <tr>
    <td height="15" bgcolor="#EFEFEF" style="padding-left:85px; padding-bottom:10px;">
		<input type="button" name="GetBak" onclick="window.location.href='../index.asp';" class="button" value="确 定" />
	</td>
  </tr>
</table>
</td>
</tr>
</table>
<%End Sub%>


<%Sub ShowErrorMsg(Title,Text)%>
<link rel="stylesheet" href="dxdiag.css" />
<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td style="border:#FF0000 1px solid;">
<table width="227" border="0" cellspacing="0">
  <tr>
    <td height="26" bgcolor="#FF6633" style="color:#FFFFFF; padding-left:10px; font-weight:bold;" class="DxdiagTitel">提示信息：</td>
  </tr>
  <tr>
    <td height="51" bgcolor="#FFFFFF" style="padding-left:10px; padding-right:10px;" class="DxdiagText"><%=Text%></td>
  </tr>
  <tr>
    <td height="15" bgcolor="#FFFFFF" style="padding-left:85px; padding-bottom:10px;">
		<input type="button" name="GetBak" onclick="window.history.go(-1);" class="button" value="返 回" />
	</td>
  </tr>
</table>
</td>
</tr>
</table>
<%End Sub%>
