<% Option Explicit %>
<% response.charset="gb2312" %>
<!--#include file="../Include/NoSqlHack.asp" -->
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/Conn2.asp" -->
<%
	Call SaveMsg()
	
	Sub SaveMsg()
		Dim UserName,ShiJian,TelPhone
		UserName=Trim(Request.Form("UserName"))
		ShiJian=Trim(Request.Form("Year"))&"-"&Trim(Request.Form("Month"))&"-"&Trim(Request.Form("Day"))
		TelPhone=Trim(Request.Form("TelPhone"))
		
		If UserName="" or isnull(UserName) or TelPhone="" or isnull(TelPhone) then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('请填写完整的信息！');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbcrlf)
			response.End()
		end if 
		
		Dim rs,sql
		Set rs=server.CreateObject("adodb.recordset")
			sql="select top 1 * from MsgData"
			rs.open sql,conn,1,3
			rs.addnew()
			rs("Msg_Name")=UserName
			rs("Msg_Time")=ShiJian
			rs("Msg_TelPhone")=TelPhone
			rs("AddTime")=now()
			rs.update()
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('留言成功！');"&vbcrlf)
				response.Write("window.location.href='Query.asp';")
			response.Write("</script>"&vbcrlf)
	End Sub
%>