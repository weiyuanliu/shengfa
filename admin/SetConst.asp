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
<TITLE>��������</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
</HEAD>
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|114,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<%
Dim Path,FileName,EditFile,FileContent,Result
Result = request.querystring("Result")
Path = "../Include"
FileName = "Const.asp"
EditFile = Server.MapPath(Path) & "\" & FileName
Dim FsoObj,FileObj,FileStreamObj
Set FsoObj = Server.CreateObject("Scripting.FileSystemObject")
Set FileObj = FsoObj.GetFile(EditFile)
if Result = "" then
	Set FileStreamObj = FileObj.OpenAsTextStream(1)
	if Not FileStreamObj.AtEndOfStream then
		FileContent = FileStreamObj.ReadAll
	else
		FileContent = ""
	end if
else
	Set FileStreamObj = FileObj.OpenAsTextStream(2)
	FileContent = Request.Form("ConstContent")
	FileStreamObj.Write FileContent
	if Err.Number <> 0 then
       response.write "<script language=javascript> alert('����ʧ�ܣ��뿽�������´��ļ��ٱ��档');location.replace('SetConst.asp');</script>"
	else
       response.write "<script language=javascript> alert('վ�㳣�������޸ĳɹ�!');location.replace('SetConst.asp');</script>"
	end if
end if
%>

<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>ϵͳ������ӣ��޸�վ��������Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap bgcolor="#EBF2F9">
	<a href="PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("�޸�����")'>�޸�����</a>	<font color="#0000FF">&nbsp;|&nbsp;</font>	<a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("��վ��Ϣ����")'>��վ��Ϣ����</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" target="mainFrame" onClick='changeAdminFlag("��Ŀ��������")'>��Ŀ��������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetConst.asp" target="mainFrame" onClick='changeAdminFlag("��������")'>��������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DataManage.asp" target="mainFrame" onClick='changeAdminFlag("���ݿ����")'>���ݿ����</a>
<font color="#0000FF">&nbsp;|&nbsp;</font><a href="ADsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("��������б�")'>�������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SpaceStat.asp" target="mainFrame" onClick='changeAdminFlag("�ռ�ͳ��")'>�ռ�ͳ��</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="../Count/InfoList.asp" target="mainFrame" onClick='changeAdminFlag("����ͳ��")'>����ͳ��</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="FriendSiteList.asp" target="mainFrame" onClick='changeAdminFlag("��������")'>��������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="HackSql.asp" target="mainFrame" onClick='changeAdminFlag("��ֹSQLע���¼")'>��ֹSQLע���¼</a>    </td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form name="ConstSetForm" action="SetConst.asp?Result=Modify" method="post">
<textarea name="ConstContent" rows="22" class="ConstSet" style="width:100%;"><% = FileContent %></textarea>
  <tr>
    <td width="10%"><input name="submitSave" type="submit" class="button" id="submitSave" value=" ���� "></td>
    <td width="90%" align="right"><font color="#FF0000">ע�⣺�����������ⵥ����"<font color="#0000FF">'</font>"��"<font color="#0000FF">&lt;%</font>"��"<font color="#0000FF">%&gt;</font>"����ȥ��������ֻ�޸��ַ�����Ҫ���ӡ�ɾ����ʹ�ûس���!</font></td>
  </tr>
</form>
</table>
</body>
</html>