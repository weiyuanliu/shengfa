<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
</head>

<body>
<%
	if AliplaySuccess("200907031705") then
		Response.Write("֧���ɹ���")
	else
		Response.Write("֧��ʧ��")
	end if
	
Function AliplaySuccess(OrderID) '֧���ɺ�Ĵ������
	if OrderID <> "" then
		Dim conn,rs,sql
		CreateConn Conn '�������Ӷ���
		CreateRs rs '������¼������
		Sql="Select State from NwebCn_Order where ProductNo='"&OrderID&"'"
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			AliplaySuccess=False
		else
			rs("State")="�����Ѹ�"
			AliplaySuccess=True
			Rs.update()
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



</body>
</html>
