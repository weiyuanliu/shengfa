<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'����������������������������������������������������������������
'����������������������������������������������������������������
'�������������������տƼ���ҵ��վ����ϵͳ��LISuo����������������  ��
'����������������������������������������������������������������
' ����Ȩ���С�qisehu.com
'
'����������������ɫ���������޹�˾
'��������������Add:�Ĵ�ʡ�ɶ��ж���·������181��13¥20/21��
'����������������������������������������������������������������
'����������������������������������������������������������������
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ���ɫ���������޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>�༭��������</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|119,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result,px
Result=request.QueryString("Result")
dim ID,SiteName,ViewFlag,LinkType,SiteFace,SiteUrl,Remark
ID=request.QueryString("ID")
call FriendSiteEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>����Ʒ�ƣ���ӣ��޸�����������ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="FriendSiteEdit.asp?Result=Add" onClick='changeAdminFlag("��ӷ���Ʒ��")'>�����������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="FriendSiteList.asp" onClick='changeAdminFlag("����Ʒ���б�")'>�鿴��������</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="FriendSiteEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">����ƶ����ܣ�</td>
        <td><input name="SiteName" type="text" class="textfield" id="SiteName" style="WIDTH: 240;" value="<%=SiteName%>">&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">����������</td>
        <td><input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag then response.write ("checked")%>></td>
      </tr>
	   <tr>
        <td height="20" align="right">����</td>
        <td><input name="px" type="text" class="textfield" id="px" style="WIDTH: 100;" value="<%=px%>">          &nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">�������ͣ�</td>
        <td><input name="LinkType" type="radio" value="1" <%if LinkType then response.write ("checked=checked")%>/>ͼƬ&nbsp;<input name="LinkType" type="radio" value="0" <%if not LinkType then response.write ("checked=checked")%>/>����</td>
      </tr>
      <tr>
        <td height="20" align="right">ǰ̨��ʾ��</td>
        <td><input name="SiteFace" type="text" class="textfield" id="SiteFace" style="WIDTH: 240;" value="<%=SiteFace%>">
        &nbsp;*&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SiteFace',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a>&nbsp;&nbsp;ͼƬ196��67&nbsp;&nbsp;���֡�8������</td>
      </tr>
      <tr>
        <td height="20" align="right">������ַ��</td>
        <td><input name="SiteUrl" type="text" class="textfield" id="SiteUrl" style="WIDTH: 490;" value="<%=SiteUrl%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">��ע˵����
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 490;"><%=Remark%></textarea></td>
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
sub FriendSiteEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("SiteName")))<2 then
      response.write ("<script language=javascript> alert('��վ����Ϊ������Ŀ�Ҳ�����2���ַ���');history.back(-1);</script>")
      response.end
    end if
    if trim(request.Form("SiteFace"))="" then
      response.write ("<script language=javascript> alert('ǰ̨��ʾΪ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
    if request.Form("LinkType")=0 then
      if StrLen(trim(request.Form("SiteFace")))>16 then
      response.write ("<script language=javascript> alert('��ѡ���""����""���ӣ����ǰ̨��ʾ���ó���8�����֣�');history.back(-1);</script>")
      response.end
      end if
    end if
    if len(trim(request.Form("SiteUrl")))<6 then
      response.write ("<script language=javascript> alert('������ַΪ������Ŀ�Ҳ�����6���ַ���');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '������վ����Ա
	  sql="select * from NwebCn_FriendSite"
      rs.open sql,conn,1,3
      rs.addnew
      rs("SiteName")=trim(Request.Form("SiteName"))
	  if isnumeric(trim(Request.Form("Px"))) then
	  rs("Px")=trim(Request.Form("Px"))
	  else
	  rs("Px")=0
	  end if
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
      rs("SiteFace")=trim(Request.Form("SiteFace"))
      rs("SiteUrl")=trim(Request.Form("SiteUrl"))
	  if Request.Form("LinkType")=1 then
        rs("LinkType")=Request.Form("LinkType")
	  else
        rs("LinkType")=0
	  end if	  
	  rs("Remark")=trim(Request.Form("Remark"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_FriendSite where ID="&ID
      rs.open sql,conn,1,3
      rs("SiteName")=trim(Request.Form("SiteName"))
	   if isnumeric(trim(Request.Form("Px"))) then
	  rs("Px")=trim(Request.Form("Px"))
	  else
	  rs("Px")=0
	  end if
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
      rs("SiteFace")=trim(Request.Form("SiteFace"))
      rs("SiteUrl")=trim(Request.Form("SiteUrl"))
	  if Request.Form("LinkType")=1 then
        rs("LinkType")=Request.Form("LinkType")
	  else
        rs("LinkType")=0
	  end if	  
	  rs("Remark")=trim(Request.Form("Remark"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭�������ӣ�');changeAdminFlag('���������б�');location.replace('FriendSiteList.asp');</script>"
  else '��ȡ��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_FriendSite where ID="& ID
      rs.open sql,conn,1,1
	  SiteName=rs("SiteName")
	  ViewFlag=rs("ViewFlag")
	  LinkType=rs("LinkType")
	  SiteFace=rs("SiteFace")
      SiteUrl=rs("SiteUrl")
      Remark=rs("Remark")
	  Px=rs("Px")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>