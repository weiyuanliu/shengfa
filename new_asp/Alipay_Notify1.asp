<%
	'���ܣ���������з�����֪ͨҳ��
	'�汾��2.0
	'���ڣ�2008-1-5
	'���ߣ�֧������˾���۲�����֧���Ŷ�
	'��ϵ��0571-26888888
	'��Ȩ��֧������˾
%>

<!--#include file="alipayto/Alipay_md5.asp"-->
<%
    key="7kyhcjza17shaxiutofguau6kjryinti"         '֧������ȫ������
    partner="2088102160488222"     '֧��������id 
 
	out_trade_no	=DelStr(Request.Form("out_trade_no"))      '��ȡ������
    total_fee		=DelStr(Request.Form("total_fee"))         '��ȡ֧�����ܼ۸�
	'�����ȡ��������������д ���� =DelStr(Request.Form("��ȡ������"))
	
'*******************�ж���Ϣ�ǲ���֧��������***********************
alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
alipayNotifyURL = alipayNotifyURL &"partner=" & partner & "&notify_id=" & request.Form("notify_id")
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
'*******************************************************************

'*******************��ȡ֧����POST����֪ͨ��Ϣ**********************
For Each varItem in Request.Form
	mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
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

'*************************����״̬���ش���*************************
If mysign=request.Form("sign") And ResponseTxt="true" Then 	
	If request.Form("trade_status") = "TRADE_FINISHED" Then 
' �����������֧�����Ĺ�����ܣ����ڷ��ص���Ϣ���治Ҫ�������жϣ���������У��ͨ���������ֵ������������Ҫ��ȡ�����ʹ�ù����Ľ��,
' ���ȡ������Ϣ������ֶ�discount��ֵ��ȡ����ֵ��������Ҹ����ŻݵĽ��� ԭ�������ܽ��=��Ҹ���صĽ��total_fee +|discount|.
		'�ڴ˴���ӣ�����ɹ�,�������ݿ����  
	 	if AliplaySuccess(out_trade_no) then '֧���ɹ��Ĵ������
			ShowMsg "��ʾ��Ϣ","��ϲ��������֧���ɹ���"
		else
			ShowErrorMsg "��ʾ��Ϣ","֧��ʧ�ܣ��뷵�أ�"	
		end if
	Else '֧��ʧ��ִ�еĳ���
		ShowErrorMsg "��ʾ��Ϣ","֧��ʧ�ܣ��뷵�أ�"			
	End If
	'Response.Write returnTxt������׷��ص�̫̬
Else
	ShowErrorMsg "��ʾ��Ϣ","���ݳ���֧��ʧ�ܣ��뷵�أ�"	
	'response.write "fail" '��ȡ���ݳ������ʾ��Ϣ
End If 
'*******************************************************************

'���������֧���ı�д����־����ô���Դ��±�ע�Ͳ��֣�������ԡ�

 'д�ı���������ԣ�����վ����Ҳ���Ըĳɴ������ݿ⣩
'TOEXCELLR=TOEXCELLR&md5str&"MD5���:"&mysign&"="&request.Form("sign")&"--ResponseTxt:"&ResponseTxt
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
	DelStr	= Replace(DelStr,"��","")
	DelStr	= Replace(DelStr,"%20","")
	DelStr	= Replace(DelStr,"--","")
	DelStr	= Replace(DelStr,"==","")
	DelStr	= Replace(DelStr,"<","")
	DelStr	= Replace(DelStr,">","")
	DelStr	= Replace(DelStr,"%","")
End Function


'////�Զ��庯��
Function AliplaySuccess(OrderID) '֧���ɺ�Ĵ������
	if OrderID <> "" then
		Dim conn,rs,sql
		CreateConn Conn '�������Ӷ���
		CreateRs rs '������¼������
		
		Sql="Select State from NwebCn_Order where ProductNo='"&OrderID&"' "
		rs.open sql,conn,1,3
		if rs("State")="" or rs("state")=null Then
		if rs.eof and rs.bof then
			AliplaySuccess=False
		else
			rs("State")="�����Ѹ�"
			AliplaySuccess=True
			Rs.update()
		end if
		end if
		CloseObject rs
		CloseObject Conn
	end if
End Function

'���� Conn����
Sub CreateConn(ByRef Conn)
	Dim ConnStr
	On error resume next
	Set Conn=Server.CreateObject("Adodb.Connection")
	ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("../Database/NwebCn_Site.asp")
	Conn.open ConnStr
	if err then
	   err.clear
	   Set Conn = Nothing
	   Response.Write "ϵͳ�������ݿ����ӳ�������'ϵͳ����>>վ�㳣������',����/Include/Const.asp�ļ�!"
	   Response.End
	end if
End Sub

'������¼������
Sub CreateRs(ByRef Object)
	Set Object=server.CreateObject("Adodb.Recordset")
End Sub

Sub CloseObject(ByRef Object)
	Object.Close()
	Set Object=Nothing
End Sub
%>
<!--��ʾ��Ϣ-->
<%Sub ShowMsg(Title,Text)%>
<link rel="stylesheet" href="dxdiag.css" />
<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td style="border:#C1C1C1 1px solid;">
<table width="227" border="0" cellspacing="0">
  <tr>
    <td height="26" bgcolor="#2A7FFF" style="color:#FFFFFF; padding-left:10px; font-weight:bold;" class="DxdiagTitel">��ʾ��Ϣ��</td>
  </tr>
  <tr>
    <td height="51" bgcolor="#EFEFEF" style="padding-left:10px; padding-right:10px;" class="DxdiagText"><%=Text%></td>
  </tr>
  <tr>
    <td height="15" bgcolor="#EFEFEF" style="padding-left:85px; padding-bottom:10px;">
		<input type="button" name="GetBak" onclick="window.location.href='../index.asp';" class="button" value="ȷ ��" />
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
    <td height="26" bgcolor="#FF6633" style="color:#FFFFFF; padding-left:10px; font-weight:bold;" class="DxdiagTitel">��ʾ��Ϣ��</td>
  </tr>
  <tr>
    <td height="51" bgcolor="#FFFFFF" style="padding-left:10px; padding-right:10px;" class="DxdiagText"><%=Text%></td>
  </tr>
  <tr>
    <td height="15" bgcolor="#FFFFFF" style="padding-left:85px; padding-bottom:10px;">
		<input type="button" name="GetBak" onclick="window.history.go(-1);" class="button" value="�� ��" />
	</td>
  </tr>
</table>
</td>
</tr>
</table>
<%End Sub%>
