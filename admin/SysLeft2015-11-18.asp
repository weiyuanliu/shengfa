<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七日科技企业网站管理系统（LISuo）　　　　　　　  ┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
' 　版权所有　qisehu.com
'
'　　程序制作　七日科技有限公司
'　　　　　　　Add:四川省成都市二环路西三段181号13楼20/21号
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
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
   Response.Write("记录不存在")
   else
   
  
  ' if rspur("qxsz")=1 then 
   leftpur=rspur("leftPurview")
   end if
  
   rspur.close
   set rspur=nothing
%>
<HTML>
<HEAD>
<TITLE>后台管理导航</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
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

var count=0;//做计数器
var limit=new Array();//用于记录当前显示的哪几个菜单
var countlimit=1;//同时打开菜单数目，可自定义

function expandIt(el) {
   obj = eval("sub" + el);
   if (obj.style.display == "none") {
      obj.style.display = "block";//显示子菜单
      if (count<countlimit) {//限制2个
         limit[count]=el;//录入数组
         count++;
      }
      else {
         eval("sub" + limit[0]).style.display = "none";
         for (i=0;i<limit.length-1;i++) {limit[i]=limit[i+1];}//数组去掉头一位，后面的往前挪一位
         limit[limit.length-1]=el;
      }
   }
   else {
      obj.style.display = "none";
      var j;
      for (i=0;i<limit.length;i++) {if (limit[i]==el) j=i;}//获取当前点击的菜单在limit数组中的位置
      for (i=j;i<limit.length-1;i++) {limit[i]=limit[i+1];}//j以后的数组全部往前挪一位
      limit[limit.length-1]=null;//删除数组最后一位
      count--;
   }
}
</script>
</HEAD>
<!--#include file="CheckAdmin.asp"-->

<BODY background="Images/SysLeft_bg.gif" onmouseover="self.status='全心全意为您打造!';return true">
    <%
	if session("GroupID") = 1 or session("GroupID")=2 then
	%>
<div id="main1" onclick=expandIt(1)     >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">页面信息</td>
    </tr>
  </table>
</div>
<div id="sub1" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"> <a href="AboutList.asp" target="mainFrame" onClick='changeAdminFlag("企业信息列表")'>页面信息列表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="AboutEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加企业信息")'>添加页面信息</a></td>
    </tr>
  </table>
</div>
<div id="main2" onclick=expandIt(2)  >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">新闻中心</td>
    </tr>
  </table>
</div>
<div id="sub2" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
   
	<tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Sort.asp?Action=Add&ParentID=0&TbS=NwebCn_NewsSort&Tb=NwebCn_News" target="mainFrame" onClick='changeAdminFlag("新闻类别")'>新闻类别</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="NewsList.asp" target="mainFrame" onClick='changeAdminFlag("新闻列表")'>新闻列表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="NewsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加新闻")'>添加新闻</a></td>
    </tr>
  </table>
</div>

<div id="main3" onclick=expandIt(3) >
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
    <tr style="cursor: hand;">
      <td width="26" ></td>
      <td class="SystemLeft">产品展示</td>
    </tr>
  </table>
</div>
<div id="sub3" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
 
	<tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Sort.asp?Action=Add&ParentID=0&TbS=NwebCn_ProductSort&Tb=NwebCn_Products" target="mainFrame" onClick='changeAdminFlag("产品类别")'>产品类别</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="ProductList.asp" target="mainFrame" onClick='changeAdminFlag("产品列表")'>产品列表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="ProductEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("添加产品")'>添加产品</a></td>
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
      <td class="SystemLeft">留言管理</td>
    </tr>
  </table>
</div>
<div id="sub9" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="MessageList.asp" target="mainFrame" onClick='changeAdminFlag("留言信息列表")'>留言信息</a></td>
    </tr>
	<%
	if session("AdminId") = 1 then
	%>
	<tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="MessageListH.asp" target="mainFrame" onClick='changeAdminFlag("留言信息列表")'>留言回收站</a></td>
    </tr>
    <%
	end if
	%>
	<tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OrderList.asp" target="mainFrame" onClick='changeAdminFlag("订单信息")'>订单信息</a></td>
    </tr>
	<%
	if session("AdminId") = 1 then
	%>
     <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OrderListH.asp" target="mainFrame" onClick='changeAdminFlag("订单信息")'>订单回收站</a></td>
    </tr>
    <%
	end if
	%>	
     <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="BlackListH.asp" target="mainFrame" onClick='changeAdminFlag("订单黑名单")'>订单黑名单</a></td>
    </tr>
    <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OrderMsg.asp" target="mainFrame" onClick='changeAdminFlag("定单留言信息")'>定单留言信息</a></td>
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
      <td class="SystemLeft">用户管理</td>
    </tr>
  </table>
</div>
<div id="sub10" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr   >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="AdminList.asp" target="mainFrame" onClick='changeAdminFlag("网站管理员")'>网站管理员</a></td>
    </tr>	
	<tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="MemGroup.asp" target="mainFrame" onClick='changeAdminFlag("管理组")'>管理组</a></td>
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
      <td class="SystemLeft">系统管理</td>
    </tr>
  </table>
</div>

<div id="sub11" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("修改密码")'>修改密码</a></td>
    </tr>
	<%
	if session("GroupID") = 1 or session("GroupID")=2 then
	%>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>网站信息设置</a></td>
    </tr>

	 <tr  > 
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="datamanage.asp" target="mainFrame" onClick='changeAdminFlag("访问日志")'>访问日志</a></td>
    </tr>
	 <tr> 
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="advlist.asp" target="mainFrame" onClick='changeAdminFlag("网络推广")'>网络推广</a></td>
    </tr>
  <tr  >
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Wenjian.asp" target="mainFrame" onClick='changeAdminFlag("文件管理")'>文件管理</a></td>
 </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="DataManage.asp" target="mainFrame" onClick='changeAdminFlag("数据库操作")'>数据库操作</a></td>
    </tr>
	<%
	end if
	%>
  </table>
</div>

<table width="170" height="24" border="0" cellpadding="0" cellspacing="0" background="Images/SysLeft_bg_click.gif">
  <tr style="cursor: hand;">
    <td width="26"></td>
    <td class="SystemLeft"><a href="javascript:AdminOut()"><font color="#ffffff">退出登录</font></a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="SysCome.asp" target="mainFrame" onClick='changeAdminFlag("后台主页")'><font color="#ffffff">后台主页</font></a></td>
  </tr>
</table>
</BODY>
</HTML>