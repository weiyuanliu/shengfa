<% Option Explicit %>
<% response.charset="gb2312" %>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/Conn2.asp" -->
<!--#include file="../Include/NoSqlHack.asp" -->
<!--#include file="../ip/cz.asp" -->
<%

dim rs,sql,message_note
set rs = server.createobject("adodb.recordset")
sql="select top 1 message_note from NwebCn_Site"
rs.open sql,conn,1,1
message_note=rs("message_note")
rs.close
set rs=nothing 

	Call SaveMsg()
	
	Sub SaveMsg()
		Dim ip,Msg_Title,Msg_Content,Linkman,ipadd,yzm
		
		
		ip=request.QueryString("ip")
if trim(ip)="" then
   IP=Request.ServerVariables("REMOTE_ADDR")
elseif ubound(split(trim(ip),"."))<>3 then
   IP=Request.ServerVariables("REMOTE_ADDR") '获取ip地址
end if
Linkman=Look_Ip("../ip/iptodata.dat",ip)  '获取ip所属地区,../ip/iptodata.dat为ip库文件位置
		
		Msg_Title=Trim(SafeRequest("Msg_Title","post"))
		Msg_Content=Trim(SafeRequest("Msg_Content","post"))
		'Linkman=Trim(SafeRequest("Linkman","post"))
		
		yzm=Trim(SafeRequest("yzm","post"))
		ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR") 
		
		if Msg_Title="" or isnull(Msg_Title) then
			response.write("<script language=javascript>"&vbcrlf)
				response.write("alert('请填写留言标题！');"&vbcrlf)
				response.write("window.history.go(-1);")
			response.write("</script>")
			response.end()
		end if
		
		if Msg_Content="" or isnull(Msg_Content) then
			response.write("<script language=javascript>"&vbcrlf)
				response.write("alert('请填写留言内容！');"&vbcrlf)
				response.write("window.history.go(-1);")
			response.write("</script>")
			response.end()
		end if
		
		if session("firstecode_left")<>yzm then
			response.write("<script language=javascript>"&vbcrlf)
				response.write("alert('请正确填写验证码！');"&vbcrlf)
				response.write("window.history.go(-1);")
			response.write("</script>")
			response.end()
		end if
		
		Dim rs,sql
		set rs=server.createobject("adodb.recordset")
		sql="select top 1 * from NwebCn_Message"
		rs.open sql,conn,1,3
		rs.addnew()
		rs("MesName")=Msg_Title
		rs("Content")=StrReplace(Msg_Content)
		rs("Linkman")=Linkman
		rs("Mobile")=ipadd
		rs("MemID")=0
		rs("AddTime")=now()
		rs("Flag")=0
		rs.update()
		rs.close()
		set rs=Nothing
		response.write("<script language=javascript>"&vbcrlf)
			response.write("alert('"&replace(message_note,vbcrlf,"\r")&"');"&vbcrlf)
			response.write("window.location.href=document.referrer;")
		response.write("</script>")
	End sub
%>