<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|314,")=0 then 
  response.write ("<script language=javascript> alert('�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�');history.back(-1);</script>")
end if
%>
<%
'========�ж��Ƿ���й���Ȩ��
Dim order_id,f
order_id=trim(Request.QueryString("order_id"))
f=Trim(Request.QueryString("f"))
if order_id="" or isnull(order_id) or not(IsNumeric(order_id))   then
	response.Write("0")
else
	dim rs,sql,sms_statas
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from NwebCn_Order where id="&order_id
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("0")
		response.End()
	else
		rs("KDFS")=f
		response.Write("1")
		rs.update()
		rs.close()
		set rs=Nothing
	end if
end if
%>
