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
<META NAME="Author" CONTENT="˳���������޹�˾" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>˳������</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<SCRIPT language=JavaScript>
function switchSysBar()
{
   if (switchPoint.innerText==3)
   {
      switchPoint.innerText=4
      document.all("frameTitle").style.display="none"
   }
   else
   {
      switchPoint.innerText=3
      document.all("frameTitle").style.display=""
   }
}
</SCRIPT>
</HEAD>
<!--#include file="CheckAdmin.asp"-->
<BODY scroll="no" topmargin="0" bottom="0" leftmargin="0" rightmargin="0">
<TABLE height="100%" cellSpacing="0" cellPadding="0" border="0" width="100%">
  <TR>
    <TD colSpan="3">
	<IFRAME 
      style="Z-INDEX:1; VISIBILITY:inherit; WIDTH:100%; HEIGHT:64" 
      name="topFrame" id="topFrame" marginWidth="0" marginHeight="0"
      src="SysTop.asp" frameBorder="0" noResize scrolling="no">	</IFRAME>	</TD>
  </TR>
  <TR>
    <TD width="170" height="100%" rowspan="2" align="middle" bgcolor="#367BDA" id="frameTitle" >
    <IFRAME 
      style="Z-INDEX:2; VISIBILITY:inherit; WIDTH:170; HEIGHT:100%" 
      name="leftFrame" id="leftFrame"  marginWidth="0" marginHeight="0"
      src="SysLeft.asp" frameBorder="0" noResize scrolling="no">
	</IFRAME>
	</TD>
    <TD width="10" height="100%" rowspan="2" align="center" bgcolor="#367BDA" onClick="switchSysBar()">
	<FONT style="FONT-SIZE: 10px; CURSOR: hand; COLOR: #ffffff; FONT-FAMILY: Webdings">
	  <SPAN id="switchPoint">3</SPAN>
	</FONT>
	</TD>
    <TD height="30">
	<iframe 
      style="Z-INDEX:3; VISIBILITY:inherit; WIDTH:100%; HEIGHT:30"
	  name="headFrame" id="mainFrame" marginwidth="16" marginheight="3"
	  src="SysHead.asp" frameborder="0"  scrolling="no">
	</iframe>
	</TD>
  </TR>
  <TR>
    <TD height="100%">
	<iframe 
      style="Z-INDEX:4; VISIBILITY:inherit; WIDTH:100%; HEIGHT:100%"
	  name="mainFrame" id="mainFrame" marginwidth="16" marginheight="16"
	  src="SysCome.asp" frameborder="0" noresize scrolling="yes">
	</iframe>
	</TD>
  </TR>
</TABLE>
</BODY>
</HTML>