<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'����������������������������������������������������������������
'����������������������������������������������������������������
'�������������������տƼ���ҵ��վ����ϵͳ��LISuo����������������  ��
'����������������������������������������������������������������
' ����Ȩ���С�qisehu.com
'
'�����������������տƼ����޹�˾
'��������������Add:�Ĵ�ʡ�ɶ��ж���·������181��13¥20/21��
'����������������������������������������������������������������
'����������������������������������������������������������������
%>
<% Option Explicit %>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>�鿴���޸ġ��ظ�����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|81,")=0 then 
  response.write ("<script language=javascript> alert('�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�');history.back(-1);</script>")
end if
%>
<%
if Instr(session("AdminPurview"),"|94,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
Dim id,Action
Action=Trim(Request.QueryString("Action"))
id=trim(Request.QueryString("ID"))
if id="" or isnull(id) or not(IsNumeric(id)) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('���ݳ����뷵�أ�');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
else
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select State,FuKuan from NwebCn_Order where id="&id
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('��¼δ�ҵ�������ʧ�ܣ�');"&vbcrlf)
			response.Write("window.history.go(-1);")
		response.Write("</script>")
		response.End()
	else
		if rs("State")="���ѷ�" then
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('���ѷ����������ظ�������');")
				response.Write("window.history.go(-1);")
			response.Write("</script>")
			response.End()
		else
			if Action="true" then
				rs("State")="�����󸶿�"
				rs("FuKuan")=true
			else
				rs("State")="���ܷ���"
				rs("FuKuan")=false
			end if
			rs.update()
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('�����ɹ���');"&vbcrlf)
				response.Write("window.location.href=document.referrer;")
			response.Write("</script>")
		end if
	end if
end if
%>
