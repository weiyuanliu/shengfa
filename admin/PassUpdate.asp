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
<TITLE>�޸�����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|111,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<body>
<%
select case request.QueryString("Action")
  case "ModifyPass"
    SaveNewPass
  case else
end select
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>ϵͳ������ӣ��޸�վ��������Ϣ</strong></font></td>
  </tr>
 <!-- <tr>
    <td height="24" align="center" nowrap bgcolor="#EBF2F9">
	<a href="PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("�޸�����")'>�޸�����</a>	<font color="#0000FF">&nbsp;|&nbsp;</font>	<a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("��վ��Ϣ����")'>��վ��Ϣ����</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" target="mainFrame" onClick='changeAdminFlag("��Ŀ��������")'>��Ŀ��������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetConst.asp" target="mainFrame" onClick='changeAdminFlag("��������")'>��������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DataManage.asp" target="mainFrame" onClick='changeAdminFlag("���ݿ����")'>���ݿ����</a>
<font color="#0000FF">&nbsp;|&nbsp;</font><a href="ADsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("��������б�")'>�������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SpaceStat.asp" target="mainFrame" onClick='changeAdminFlag("�ռ�ͳ��")'>�ռ�ͳ��</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="../Count/InfoList.asp" target="mainFrame" onClick='changeAdminFlag("����ͳ��")'>����ͳ��</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="FriendSiteList.asp" target="mainFrame" onClick='changeAdminFlag("��������")'>��������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="HackSql.asp" target="mainFrame" onClick='changeAdminFlag("��ֹSQLע���¼")'>��ֹSQLע���¼</a>    </td>
  </tr>-->
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="PassUpdate.asp?Action=ModifyPass&LoginName=<%=session("AdminName")%>" >
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="220" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;¼&nbsp;����</td>
        <td><input name="AdminName" type="text" class="textfield" id="AdminName" style="WIDTH: 120;" value="<%=session("AdminName")%>" maxlength="16" readonly>&nbsp;3-10λ�ַ��������޸�</td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;��&nbsp;�룺</td>
        <td><input name="NewPassword" type="password" class="textfield" id="NewPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*&nbsp;ע����ĸ��Сд</td>
      </tr>
      <tr>
        <td height="20" align="right">ȷ�����룺</td>
        <td><input name="vNewPassword" type="password" class="textfield" id="vNewPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*</td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="����" style="WIDTH: 60;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</body>
</html>
<%
function SaveNewPass()
  dim LoginName,rs,sql 
  LoginName=request.QueryString("LoginName")
  set rs = server.createobject("adodb.recordset")
  sql="select * from NwebCn_Admin where AdminName='"&LoginName&"'"
  rs.open sql,conn,1,3
  if rs.bof and rs.eof then
    response.write "��ȡ���ݿ��¼����"
    response.end
  else
	if len(trim(Request.Form("NewPassword")))<6 or len(trim(Request.Form("NewPassword")))>20  then
      response.write "<script language=javascript> alert('����Ա���������ַ���Ϊ6-20λ��');history.back(-1);</script>"
      response.end
    end if
	if Request.Form("NewPassword")<>Request.Form("vNewPassword") then 
      response.write "<script language=javascript> alert('������������벻һ����');history.back(-1);</script>"
      response.end
	end if
	rs("Password")=Md5(Request.Form("NewPassword"))  
    rs.update
    rs.close
    set rs=nothing 
  end if
  response.write "��Ĺ��������ѳɹ��޸ģ����μ�[&nbsp;<font color='red'>"&trim(Request.Form("NewPassword"))&"</font>&nbsp;]��"
  response.end
end function
%>
