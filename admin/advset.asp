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
dim ID,ADS_Name,ADS_Link,AddTime
ID=request.QueryString("ID")
call ADVEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��棺��ӣ��޸Ĺ����ص�����</strong></font></td>
  </tr>
  <tr>
        <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="advset.asp?Result=Add" onClick='changeAdminFlag("��ӹ��")'>��ӹ��</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="advlist.asp" onClick='changeAdminFlag("����б�")'>�鿴���</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="advset.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">������ƣ�</td>
        <td><input name="ADS_Name" type="text" class="textfield" id="ADS_Name" style="WIDTH: 240;" value="<%=ADS_Name%>">&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">������ַ��</td>
        <td><input name="ADS_Link" type="text" class="textfield" id="ADS_Link" style="WIDTH: 490;" value="<%=ADS_Link%>">&nbsp;*</td>
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
function ADVEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("ADS_Name")))<2 then
      response.write ("<script language=javascript> alert('�������Ϊ������Ŀ�Ҳ�����2���ַ���');history.back(-1);</script>")
      response.end
    end if
    if len(trim(request.Form("ADS_Link")))<2 then
      response.write ("<script language=javascript> alert('����ַΪ������Ŀ�Ҳ�����10���ַ���');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then 
	  sql="select * from NwebCn_Ads_effect"
      rs.open sql,conn,1,3
      rs.addnew
      rs("ADS_Name")=trim(Request.Form("ADS_Name"))
      rs("ADS_Link")=trim(Request.Form("ADS_Link"))
      rs("AddTime")=now
	end if
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Ads_effect where ID="&ID
      rs.open sql,conn,1,3
      rs("ADS_Name")=trim(Request.Form("ADS_Name"))
      rs("ADS_Link")=trim(Request.Form("ADS_Link"))
	end if
	  rs.update
	  rs.close
      set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭��棡');changeAdminFlag('����б�');location.replace('advlist.asp');</script>"
  else '��ȡ��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Ads_effect where ID="& ID
      rs.open sql,conn,1,1
	  ADS_Name=rs("ADS_Name")
	  ADS_Link=rs("ADS_Link")
	  rs.close
      set rs=nothing 
	end if
  end if
end function
%>