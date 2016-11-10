<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
<%
	if AliplaySuccess("200907031705") then
		Response.Write("支付成功！")
	else
		Response.Write("支付失败")
	end if
	
Function AliplaySuccess(OrderID) '支付成后的处理程序
	if OrderID <> "" then
		Dim conn,rs,sql
		CreateConn Conn '建立链接对像
		CreateRs rs '建立记录集对像
		Sql="Select State from NwebCn_Order where ProductNo='"&OrderID&"'"
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			AliplaySuccess=False
		else
			rs("State")="货款已付"
			AliplaySuccess=True
			Rs.update()
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



</body>
</html>
