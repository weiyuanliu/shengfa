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
if Instr(session("AdminPurview"),"|314,")=0 then 
  response.write ("<script language=javascript> alert('�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�');history.back(-1);</script>")
end if
%>
<%
'========�ж��Ƿ���й���Ȩ��
Dim id,States
id=trim(Request.QueryString("ID"))
States=Trim(Request.QueryString("State"))
if id="" or isnull(id) or not(IsNumeric(id))   then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('���ݳ����뷵�أ�');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
else
	dim rs,sql,sms_states
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from NwebCn_Order where id="&id
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('�Բ��𣬼�¼δ�ҵ����뷵�أ�');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>")
		response.End()
	else
		'if instr(States,"Ǯ���ѷ�")>0 then
		rs("FaHuoTime")=Now()
		'else
			'if Not(rs("FaHuoTime")<>"") then
				'rs("FaHuoTime") = Now()
			'end if
		'end if
		sms_states = rs("sms_states")
		if (States = "Ǯ���ѷ�" or States = "�Ѿ�����") and sms_states=0 then
		  call sendSms(2,rs("Linkman"),rs("Tel"))
		  rs("sms_states")=1
		  response.Write("״̬��"& States )
		  'response.End()
		  else
		  rs("sms_states")=0
		end if
		'rs("sms_states")=0
		rs("State")=States
		rs.update()
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
	'		response.Write("alert('�����ɹ���');"&vbcrlf)
			response.Write("window.location.href=document.referrer;")
		response.Write("</script>"&vbcrlf)
	end if
end if
%>
