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
<TITLE>�༭���</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,ADsName,ViewFlag,Content
dim ADsWidth,ADsHeight
ID=request.QueryString("ID")
call ADsEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>������棺��ӣ��޸ĵ��������ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="ADsEdit.asp?Result=Add" onClick='changeAdminFlag("��ӵ������")'>��ӵ������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="ADsList.asp" onClick='changeAdminFlag("��������б�")'>�鿴�������</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="ADsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">�����⣺</td>
        <td><input name="ADsName" type="text" class="textfield" id="ADsName" style="WIDTH: 240;" value="<%=ADsName%>" maxlength="100">&nbsp;������<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag then response.write ("checked")%>>
&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">�����ߴ磺</td>
        <td><input name="ADsWidth" type="text" class="textfield" id="ADsWidth" style="WIDTH: 60;" value="<%=ADsWidth%>" maxlength="4" onKeyDown="if(event.keyCode==13)event.returnValue=false" onChange="if(/\D/.test(this.value)){alert('��Ⱥ͸߶�ֻ������������');this.value='150';}">&nbsp;�����&nbsp;<input name="ADsHeight" type="text" class="textfield" id="ADsHeight" style="WIDTH: 60;" value="<%=ADsHeight%>" maxlength="4" onKeyDown="if(event.keyCode==13)event.returnValue=false" onChange="if(/\D/.test(this.value)){alert('��Ⱥ͸߶�ֻ������������');this.value='100';}">&nbsp;*&nbsp;����150��100����</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">�������ݣ�<br>
		  <img title="���������ӻ��鿴���༭����..." src="Images/Edit.gif" width="51" height="20" style="cursor:hand" onClick="OpenDialog('../Editor/EditorDialog.html?lnk=Content&file=Editor_1.html',800,520);">
        <td><textarea name="Content" rows="12" class="textfield" id="Content" style="WIDTH: 86%;" readonly><%=Content%></textarea></td>
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
sub ADsEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("ADsName")))<3 then
      response.write ("<script language=javascript> alert('������Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
	if trim(request.Form("ADsWidth"))="" or trim(request.Form("ADsHeight"))="" then
      response.write ("<script language=javascript> alert('������������Ϊ150��100�������ϣ�');history.back(-1);</script>")
      response.end
	end if
	if trim(request.Form("ADsWidth"))<150 or trim(request.Form("ADsHeight"))<100 then
      response.write ("<script language=javascript> alert('������������Ϊ150��100�������ϣ�');history.back(-1);</script>")
      response.end
	end if
    if Result="Add" then '������վ����Ա
	  sql="select * from NwebCn_ADs"
      rs.open sql,conn,1,3
      rs.addnew
      rs("ADsName")=trim(Request.Form("ADsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
	  rs("ADsWidth")=trim(Request.Form("ADsWidth"))
	  rs("ADsHeight")=trim(Request.Form("ADsHeight"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_ADs where ID="&ID
      rs.open sql,conn,1,3
      rs("ADsName")=trim(Request.Form("ADsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
	  rs("ADsWidth")=trim(Request.Form("ADsWidth"))
	  rs("ADsHeight")=trim(Request.Form("ADsHeight"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭������棡');changeAdminFlag('��������б�');location.replace('ADsList.asp');</script>"
  else '��ȡ����Ա��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_ADs where ID="& ID
      rs.open sql,conn,1,1
	  ADsName=rs("ADsName")
	  ViewFlag=rs("ViewFlag")
	  ADsWidth=rs("ADsWidth")
	  ADsHeight=rs("ADsHeight")
      Content=rs("Content")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>