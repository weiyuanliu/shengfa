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
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>�༭����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|113,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,NavName,ViewFlag,NavUrl,Remark
ID=request.QueryString("ID")
call NavEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>������Ŀ����ӣ��޸ĵ�����Ŀ��ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="NavigationEdit.asp?Result=Add" onClick='changeAdminFlag("��ӵ�����Ŀ")'>��ӵ�����Ŀ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" onClick='changeAdminFlag("������Ŀ�б�")'>�鿴������Ŀ</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="NavigationEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">�������ƣ�</td>
        <td><input name="NavName" type="text" class="textfield" id="NavName" style="WIDTH: 240;" value="<%=NavName%>" maxlength="100">&nbsp;������<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag  or Result="Add" then response.write ("checked")%>>&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">������ַ��</td>
        <td><input name="NavUrl" type="text" class="textfield" id="NavUrl" style="WIDTH: 480;" value="<%=NavUrl%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">��ע˵����
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 480;"><%=Remark%></textarea></td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="����" style="WIDTH: 80;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub NavEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("NavName")))<3 then
      response.write ("<script language=javascript> alert('��������Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '������վ����Ա
	  sql="select * from NwebCn_Navigation"
      rs.open sql,conn,1,3
      rs.addnew
      rs("NavName")=trim(Request.Form("NavName"))
      rs("NavUrl")=trim(Request.Form("NavUrl"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	  rs("Sequence")=99
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Navigation where ID="&ID
      rs.open sql,conn,1,3
      rs("NavName")=trim(Request.Form("NavName"))
      rs("NavUrl")=trim(Request.Form("NavUrl"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭������Ŀ��');changeAdminFlag('������Ŀ�б�');location.replace('NavigationList.asp');</script>"
  else '��ȡ����Ա��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Navigation where ID="& ID
      rs.open sql,conn,1,1
	  NavName=rs("NavName")
	  ViewFlag=rs("ViewFlag")
      Remark=rs("Remark")
      NavUrl=rs("NavUrl")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>