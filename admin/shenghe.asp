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
<TITLE>��ˡ��޸ġ��ظ�����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|92,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<%Call Set_TuiJian()%>
<%
	Sub Set_TuiJian()
		Dim ID,Rs,Sql
		ID=Trim(Request.QueryString("ID"))
		if ID="" or Isnull(ID) or Not(IsNumeric(ID)) then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('�Բ������ݳ����뷵�أ�');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbCrlf)
			response.End()
			exit sub
		end if
		
		Set Rs=server.CreateObject("adodb.recordset")
		sql="select ViewFlag from NwebCn_Message where id="&id
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('��¼δ�ҵ�������ʧ�ܣ�');"&vbcrlf)
				response.Write("window.history.go(-1);")
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		else
			if trim(Request.QueryString("Action"))="true" then 
				rs("ViewFlag")=1
			else
				rs("ViewFlag")=0
			end if
			rs.update()
		end if
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('�����ɹ���');")
			response.Write("window.location.href=document.referrer;")
		response.Write("</script>"&vbcrlf)
	End Sub
%>
</BODY>
</HTML>
