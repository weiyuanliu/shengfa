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
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>��վ�ռ�ͳ��</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|116,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<body>
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
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td width="140" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ͳ����Ŀ</strong></font></td>
    <td height="24" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">·��</font></strong></td>
    <td width="66" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>��С</strong></font></td>
    <td width="46" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>�ٷֱ�</strong></font></td>
    <td width="260" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ͼʾ</strong></font></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >ϵͳ��ռ�ÿռ�</td>
    <td ><%=SysRootDir%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir)),,-1)%></td>
    <td ><img src="images/bar.gif" width="260" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >����ͳ���ļ���</td>
    <td ><%=SysRootDir&"Count"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Count"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Count")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Count"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >���ݿ����ļ���</td>
    <td ><%=SysRootDir&"Database"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Database"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Database")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Database"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >��վ�������ݿ�</td>
    <td ><%=SiteDataPath%></td>
    <td ><%=SizeInfo(SizeByte(SiteDataPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SiteDataPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SiteDataPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >��վ�������ݿ�</td>
    <td ><%=SiteDataBakPath%></td>
    <td ><%=SizeInfo(SizeByte(SiteDataBakPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SiteDataBakPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SiteDataBakPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�����������ݿ�</td>
    <td ><%=StatDataPath%></td>
    <td ><%=SizeInfo(SizeByte(StatDataPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(StatDataPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(StatDataPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�����������ݿ�</td>
    <td ><%=StatDataBakPath%></td>
    <td ><%=SizeInfo(SizeByte(StatDataBakPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(StatDataBakPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(StatDataBakPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >���߱༭���ļ���</td>
    <td ><%=SysRootDir&"Editor"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Editor"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Editor")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Editor"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >ǰ̨��ҳ�ļ���</td>
    <td ><%=SysRootDir&"Html"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Html"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Html")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Html"))*260%>" height="8"></td>
  </tr>  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >ǰ̨ͼƬ�ļ���</td>
    <td ><%=SysRootDir&"Images"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Images"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Images")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Images"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >��վ�����ļ�</td>
    <td ><%=SysRootDir&"Include"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Include"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Include")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Include"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >��վǰ̨JS�ű��ļ�</td>
    <td ><%=SysRootDir&"Script"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Script"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Script")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Script"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�����̨�ļ���</td>
    <td ><%=SysRootDir&"System"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"System"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"System")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"System"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�����̨ͼƬ�ļ���</td>
    <td ><%=SysRootDir&"System/Images"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"System/Images"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"System/Images")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"System/Images"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�ļ��ϴ�����Ŀ¼</td>
    <td ><%=SysRootDir&"Upload"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >ͼƬ�ļ��ϴ�����Ŀ¼</td>
    <td ><%=SysRootDir&"Upload/PicFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/PicFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/PicFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/PicFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�����ļ��ϴ�����Ŀ¼</td>
    <td ><%=SysRootDir&"Upload/DownFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/DownFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/DownFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/DownFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�����ļ��ϴ�����Ŀ¼</td>
    <td ><%=SysRootDir&"Upload/OtherFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/OtherFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/OtherFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/OtherFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >�༭���ϴ��ļ�����Ŀ¼</td>
    <td ><%=SysRootDir&"Upload/EditorFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/EditorFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/EditorFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/EditorFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >��ʱ�ļ�</td>
    <td ><%=SysRootDir&"Temp"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Temp"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Temp")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Temp"))*260%>" height="8"></td>
  </tr>
</table>
</body>
</html>

<%
'============
function SizeByte(Path)
  dim fso
  set fso=server.createobject("scripting.filesystemobject") 
  Path=server.mappath(Path) 
  if fso.FileExists(Path) then	
 	SizeByte=fso.getfile(Path).size
  elseif fso.FolderExists(Path) then
 	SizeByte=fso.getfolder(Path).size
  else
    SizeByte="PathError"
  end if
end function
'============
function SizeInfo(SizeByte)
  if SizeByte="PathError" then
    SizeInfo="<font color='red'>δ�ҵ�</font>"
  else
    if SizeByte>=1024*1024*1024 then
      SizeInfo=round(SizeByte/1024/1024/1024,2) & " GB"		
    elseif 1024*1024<=SizeByte and SizeByte<1024*1024*1024 then
      SizeInfo=round(SizeByte/1024/1024,2) & " MB"		
    elseif 1024<=SizeByte and SizeByte<1024*1024 then
      SizeInfo=round(SizeByte/1024,2) & " KB"	
    else
      SizeInfo=SizeByte & " Byte" 
    end if
  end if
end function
'============
function Percent(SizeByte)
  dim fso,SysSizeByte
  set fso=server.createobject("scripting.filesystemobject") 
  if fso.FolderExists(server.mappath(SysRootDir)) and SizeByte<>"PathError"  then
 	SysSizeByte=fso.getfolder(server.mappath(SysRootDir)).size
    Percent=SizeByte/SysSizeByte
  else
    Percent=0
	exit function
  end if
end function
'============
%>