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
<TITLE>����Ա�б�</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
Dim Action,ID
ID=Trim(Request("ID"))
if ID="" or isnull(ID) or Not(IsNumeric(ID)) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('���ݳ����뷵�أ�');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>"&vbcrlf)
	response.End()
end if
Action=Trim(Request.QueryString("Action"))
if Action="AddProvince" then Call SaveProvince()
Dim Content,AddTime,Px
Call FuZhi()

Sub FuZhi()
	Dim Rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from Province where id="&ID
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('��¼δ�ҵ����뷵�أ�');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	else
		Content=rs("Content")
		AddTime=rs("AddTime")
		Px=rs("Px")
	end if
	rs.close()
	set rs=Nothing
End Sub
%>

<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>���һ��ʡ����Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AddProvince.asp?Result=Add" onClick='changeAdminFlag("���ʡ����Ϣ")'>���ʡ����Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="Province.asp" onClick='changeAdminFlag("�鿴����ʡ����Ϣ")'>�鿴����ʡ����Ϣ</a></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9">
    <table width="78%" border="0" cellpadding="5" cellspacing="0">
       <form name="EditProvince" id="EditProvince" method="post" action="EditProvince.asp?Action=AddProvince" onSubmit="return CheckValue();">
      <tr>
        <td width="21%" height="25" align="right">ʡ�����ƣ�</td>
        <td width="79%" height="25"><input name="Content" type="text" id="Content" value="<%=Content%>">��
          *����</td>
      </tr>
      <tr>
        <td height="25" align="right">���ʱ�䣺</td>
        <td height="25"><input name="AddTime" type="text" id="AddTime" value="<%=AddTime%>"></td>
      </tr>
      <tr>
        <td height="25" align="right">����˳��</td>
        <td height="25"><input name="Px" type="text" id="Px" value="<%=Px%>"> ��
          *����д���֣�ԽС����Խǰ</td>
      </tr>
      
      <tr>
        <td height="40" align="right">&nbsp;</td>
        <td height="40"><label>
          <input type="hidden" name="ID" value="<%=ID%>">
          <input type="submit" name="button" id="button" value="�� ��" style="margin-right:10px;">
          <input type="button" name="button2" id="button2" value="�� ��" onClick="window.history.go(-1);">
        </label></td>
      </tr>
        </form>
    </table>
  
    </td>    
  </tr>
</table>
</body>
</html>
<%
sub SaveProvince()
	dim Content,AddTime,Px,ID
	ID=Trim(Request.Form("ID"))
	Content=Trim(Request.Form("Content"))
	AddTime=Trim(Request.Form("AddTime"))
	Px=Trim(Request.Form("Px"))
	
	if Content="" or isnull(Content) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('����дʡ����Ϣ��');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	end if
	
	if Px="" or not(IsNumeric(Px)) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('����д��Ч��������Ϣ��');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	end if
	
	Dim rs,sql
	set rs=server.CreateObject("Adodb.Recordset")
	
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		sql="select * from Province where Content='"&Content&"'"
	else
		sql="select * from Province where Content='"&Content&"' and id not in("&id&")"
	end if
	rs.open sql,conn,1,1
	
	if not rs.eof and not rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script langauge=javascript>"&vbcrlf)
			response.Write("alert('�Բ��𣬲����ظ���');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	end if
	rs.close()
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		Sql="Select top 1 * from Province"
	else
		Sql="Select top 1 * from Province where id="&ID
	end if
	rs.open sql,conn,1,3
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		rs.addnew()
		rs("Content")=Content
		if AddTime<>"" then
			rs("AddTime")=AddTime
		else
			rs("AddTime")=Now()
		end if
		rs("Px")=Px
		rs.update()
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('�����ɹ���');"&vbcrlf)
			response.Write("window.location.href=document.referrer;")
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	else
		if rs.eof and rs.bof then
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('��¼δ�ҵ�������ʧ�ܣ�');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		else
			rs("Content")=Content
			if AddTime<>"" then
				rs("AddTime")=AddTime
			else
				rs("AddTime")=Now()
			end if
			rs("Px")=Px
			rs.update()
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('���¼�¼�ɹ���');"&vbcrlf)
				response.Write("window.location.href='Province.asp';")
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		end if
	end if
end sub
%>

<script language="javascript">
<!--
	function CheckValue()
	{
		var Content,Px;
		Content=document.getElementById("Content");
		Px=document.getElementById("Px");
		
		if(Content.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����дʡ�����ƣ�");
			Content.focus();
			return false;
		}
		
		if((Px.value).search("^-?\\d+(\\.\\d+)?$")!=0)
		{
			alert("����д��Ч�����֣�");
			Px.select();
			return false;
		}
		return true;
	}

-->
</script>
