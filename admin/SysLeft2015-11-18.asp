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
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim rspur,sqlpur,leftpur
   set Rspur=server.CreateObject("Adodb.recordset")
   sqlpur="select top 1 * from Purview"
   rspur.open sqlpur,conn,1,3
   if rspur.bof and rspur.eof then 
   Response.Write("��¼������")
   else
   
  
  ' if rspur("qxsz")=1 then 
   leftpur=rspur("leftPurview")
   end if
  
   rspur.close
   set rspur=nothing
%>
<HTML>
<HEAD>
<TITLE>��̨������</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script>
function closewin() {
   if (opener!=null && !opener.closed) {
      opener.window.newwin=null;
      opener.openbutton.disabled=false;
      opener.closebutton.disabled=true;
   }
}

var count=0;//��������
var limit=new Array();//���ڼ�¼��ǰ��ʾ���ļ����˵�
var countlimit=1;//ͬʱ�򿪲˵���Ŀ�����Զ���

function expandIt(el) {
   obj = eval("sub" + el);
   if (obj.style.display == "none") {
      obj.style.display = "block";//��ʾ�Ӳ˵�
      if (count<countlimit) {//����2��
         limit[count]=el;//¼������
         count++;
      }
      else {
         eval("sub" + limit[0]).style.display = "none";
         for (i=0;i<limit.length-1;i++) {limit[i]=limit[i+1];}//����ȥ��ͷһλ���������ǰŲһλ
         limit[limit.length-1]=el;
      }
   }
   else {
      obj.style.display = "none";
      var j;
      for (i=0;i<limit.length;i++) {if (limit[i]==el) j=i;}//��ȡ��ǰ����Ĳ˵���limit�����е�λ��
      for (i=j;i<limit.length-1;i++) {limit[i]=limit[i+1];}//j�Ժ������ȫ����ǰŲһλ
      limit[limit.length-1]=null;//ɾ���������һλ
      count--;
   }
}
</script>
</HEAD>
<!--#include file="CheckAdmin.asp"-->

<BODY background="Images/SysLeft_bg.gif" onmouseover="self.status='ȫ��ȫ��Ϊ������!';return true">
    <%
	if session("GroupID") = 1 or session("GroupID")=2 then
	%>
<div id="main1" onclick=expandIt(1)     >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">ҳ����Ϣ</td>
    </tr>
  </table>
</div>
<div id="sub1" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"> <a href="AboutList.asp" target="mainFrame" onClick='changeAdminFlag("��ҵ��Ϣ�б�")'>ҳ����Ϣ�б�</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="AboutEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("�����ҵ��Ϣ")'>���ҳ����Ϣ</a></td>
    </tr>
  </table>
</div>
<div id="main2" onclick=expandIt(2)  >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">��������</td>
    </tr>
  </table>
</div>
<div id="sub2" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
   
	<tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Sort.asp?Action=Add&ParentID=0&TbS=NwebCn_NewsSort&Tb=NwebCn_News" target="mainFrame" onClick='changeAdminFlag("�������")'>�������</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="NewsList.asp" target="mainFrame" onClick='changeAdminFlag("�����б�")'>�����б�</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="NewsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("�������")'>�������</a></td>
    </tr>
  </table>
</div>

<div id="main3" onclick=expandIt(3) >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">��Ʒչʾ</td>
    </tr>
  </table>
</div>
<div id="sub3" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
 
	<tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Sort.asp?Action=Add&ParentID=0&TbS=NwebCn_ProductSort&Tb=NwebCn_Products" target="mainFrame" onClick='changeAdminFlag("��Ʒ���")'>��Ʒ���</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="ProductList.asp" target="mainFrame" onClick='changeAdminFlag("��Ʒ�б�")'>��Ʒ�б�</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="ProductEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("��Ӳ�Ʒ")'>��Ӳ�Ʒ</a></td>
    </tr>
  </table>
</div>
	<%
	end if
	%>
<div id="main9" onclick=expandIt(9)  >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">���Թ���</td>
    </tr>
  </table>
</div>
<div id="sub9" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="MessageList.asp" target="mainFrame" onClick='changeAdminFlag("������Ϣ�б�")'>������Ϣ</a></td>
    </tr>
	<%
	if session("AdminId") = 1 then
	%>
	<tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="MessageListH.asp" target="mainFrame" onClick='changeAdminFlag("������Ϣ�б�")'>���Ի���վ</a></td>
    </tr>
    <%
	end if
	%>
	<tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OrderList.asp" target="mainFrame" onClick='changeAdminFlag("������Ϣ")'>������Ϣ</a></td>
    </tr>
	<%
	if session("AdminId") = 1 then
	%>
     <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OrderListH.asp" target="mainFrame" onClick='changeAdminFlag("������Ϣ")'>��������վ</a></td>
    </tr>
    <%
	end if
	%>	
     <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="BlackListH.asp" target="mainFrame" onClick='changeAdminFlag("����������")'>����������</a></td>
    </tr>
    <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OrderMsg.asp" target="mainFrame" onClick='changeAdminFlag("����������Ϣ")'>����������Ϣ</a></td>
    </tr>
  </table>
</div>
    <%
	if session("GroupID") = 1 then
	%>
<div id="main10" onclick=expandIt(10)  >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">�û�����</td>
    </tr>
  </table>
</div>
<div id="sub10" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr   >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="AdminList.asp" target="mainFrame" onClick='changeAdminFlag("��վ����Ա")'>��վ����Ա</a></td>
    </tr>	
	<tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="MemGroup.asp" target="mainFrame" onClick='changeAdminFlag("������")'>������</a></td>
    </tr>
	<%
	end if
	%>
  </table>
</div>

<div id="main11" onclick=expandIt(11)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">ϵͳ����</td>
    </tr>
  </table>
</div>

<div id="sub11" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("�޸�����")'>�޸�����</a></td>
    </tr>
	<%
	if session("GroupID") = 1 or session("GroupID")=2 then
	%>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("��վ��Ϣ����")'>��վ��Ϣ����</a></td>
    </tr>

	 <tr  > 
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="datamanage.asp" target="mainFrame" onClick='changeAdminFlag("������־")'>������־</a></td>
    </tr>
	 <tr> 
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="advlist.asp" target="mainFrame" onClick='changeAdminFlag("�����ƹ�")'>�����ƹ�</a></td>
    </tr>
  <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Wenjian.asp" target="mainFrame" onClick='changeAdminFlag("�ļ�����")'>�ļ�����</a></td>
 </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="DataManage.asp" target="mainFrame" onClick='changeAdminFlag("���ݿ����")'>���ݿ����</a></td>
    </tr>
	<%
	end if
	%>
  </table>
</div>

<table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
  <tr style="cursor: hand;">
    <td width="26"></td>
    <td class="SystemLeft"><a href="javascript:AdminOut()"><font color="#ffffff">�˳���¼</font></a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="SysCome.asp" target="mainFrame" onClick='changeAdminFlag("��̨��ҳ")'><font color="#ffffff">��̨��ҳ</font></a></td>
  </tr>
</table>
</BODY>
</HTML>