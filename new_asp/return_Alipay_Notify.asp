<%
	'���ܣ���������ת��ҳ��
	'�汾��2.0
	'���ڣ�2008-1-5
	'���ߣ�֧������˾���۲�����֧���Ŷ�
	'��ϵ��0571-26888888
	'��Ȩ��֧������˾
%>

<!--#include file="alipayto/Alipay_md5.asp"-->
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->

<%
	  partner			="2088402551356533"	 '��д��Ӧ֧�����˻��ĺ��������ID
	  key			    ="dinrqtpwtcai6wzv4iy8qby016hb67uo"	 '��д��Ӧ֧�����ʻ��İ�ȫУ����

	out_trade_no	= DelStr(Request("out_trade_no"))  '��ȡ������
    total_fee		= DelStr(Request("total_fee"))     '��ȡ֧�����ܼ۸�
	'�����ȡ��������������д ���� =DelStr(Request.Form("��ȡ������"))

'**********************�ж���Ϣ�ǲ���֧��������********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    '��ʱע����������2013-8-29
    'Retrieval.send()
    'ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*******************************************************************

'*******��ȡ֧����GET����֪ͨ��Ϣ,�ж���Ϣ�ǲ��Ǳ��޸Ĺ�************
For Each varItem in Request.QueryString
	mystr=varItem&"="&Request(varItem)&"^"&mystr
Next 
If mystr<>"" Then 
	mystr=Left(mystr,Len(mystr)-1)
End If 
mystr = SPLIT(mystr, "^")
Count=ubound(mystr)
'�Բ�������
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
'����md5ժҪ�ַ���
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
	response.write "Ҳ����֧���ɹ������ǵ���վȴû�м�¼�������ݣ�����ϵ�ͷ�ȷ��!"          '�������ָ������Ҫ��ʾ������
End If 


Function DelStr(Str)
	If IsNull(Str) Or IsEmpty(Str) Then
		Str	= ""
	End If
	DelStr	= Replace(Str,";","")
	DelStr	= Replace(DelStr,"'","")
	DelStr	= Replace(DelStr,"&","")
	DelStr	= Replace(DelStr," ","")
	DelStr	= Replace(DelStr,"��","")
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
  rs("State")="�����Ѹ�"
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
  Text=replace(Text,"{�������}",out_trade_no)
  Text=replace(Text,"{֧�����}",total_fee)
  response.Write Text
  rs("ClickNumber")=rs("ClickNumber")+1
  rs.update
 end if
 rs.close
 set rs=nothing
End Function

%>