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
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>网站空间统计</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|116,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>系统管理：添加，修改站点的相关信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap bgcolor="#EBF2F9">
	<a href="PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("修改密码")'>修改密码</a>	<font color="#0000FF">&nbsp;|&nbsp;</font>	<a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>网站信息设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" target="mainFrame" onClick='changeAdminFlag("栏目导航设置")'>栏目导航设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetConst.asp" target="mainFrame" onClick='changeAdminFlag("常量设置")'>常量设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DataManage.asp" target="mainFrame" onClick='changeAdminFlag("数据库操作")'>数据库操作</a>
<font color="#0000FF">&nbsp;|&nbsp;</font><a href="ADsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("弹窗广告列表")'>弹窗广告</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SpaceStat.asp" target="mainFrame" onClick='changeAdminFlag("空间统计")'>空间统计</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="../Count/InfoList.asp" target="mainFrame" onClick='changeAdminFlag("访问统计")'>访问统计</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="FriendSiteList.asp" target="mainFrame" onClick='changeAdminFlag("友情链接")'>友情链接</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="HackSql.asp" target="mainFrame" onClick='changeAdminFlag("阻止SQL注入记录")'>阻止SQL注入记录</a>    </td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td width="140" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>统计栏目</strong></font></td>
    <td height="24" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">路径</font></strong></td>
    <td width="66" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>大小</strong></font></td>
    <td width="46" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>百分比</strong></font></td>
    <td width="260" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>图示</strong></font></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >系统总占用空间</td>
    <td ><%=SysRootDir%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir)),,-1)%></td>
    <td ><img src="images/bar.gif" width="260" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >流量统计文件夹</td>
    <td ><%=SysRootDir&"Count"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Count"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Count")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Count"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >数据库存放文件夹</td>
    <td ><%=SysRootDir&"Database"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Database"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Database")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Database"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >网站常规数据库</td>
    <td ><%=SiteDataPath%></td>
    <td ><%=SizeInfo(SizeByte(SiteDataPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SiteDataPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SiteDataPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >网站备份数据库</td>
    <td ><%=SiteDataBakPath%></td>
    <td ><%=SizeInfo(SizeByte(SiteDataBakPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SiteDataBakPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SiteDataBakPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >流量常规数据库</td>
    <td ><%=StatDataPath%></td>
    <td ><%=SizeInfo(SizeByte(StatDataPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(StatDataPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(StatDataPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >流量备份数据库</td>
    <td ><%=StatDataBakPath%></td>
    <td ><%=SizeInfo(SizeByte(StatDataBakPath))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(StatDataBakPath)),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(StatDataBakPath))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >在线编辑器文件夹</td>
    <td ><%=SysRootDir&"Editor"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Editor"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Editor")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Editor"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >前台网页文件夹</td>
    <td ><%=SysRootDir&"Html"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Html"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Html")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Html"))*260%>" height="8"></td>
  </tr>  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >前台图片文件夹</td>
    <td ><%=SysRootDir&"Images"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Images"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Images")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Images"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >网站包含文件</td>
    <td ><%=SysRootDir&"Include"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Include"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Include")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Include"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >网站前台JS脚本文件</td>
    <td ><%=SysRootDir&"Script"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Script"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Script")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Script"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >管理后台文件夹</td>
    <td ><%=SysRootDir&"System"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"System"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"System")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"System"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >管理后台图片文件夹</td>
    <td ><%=SysRootDir&"System/Images"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"System/Images"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"System/Images")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"System/Images"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >文件上传保存目录</td>
    <td ><%=SysRootDir&"Upload"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >图片文件上传保存目录</td>
    <td ><%=SysRootDir&"Upload/PicFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/PicFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/PicFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/PicFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >下载文件上传保存目录</td>
    <td ><%=SysRootDir&"Upload/DownFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/DownFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/DownFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/DownFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >其他文件上传保存目录</td>
    <td ><%=SysRootDir&"Upload/OtherFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/OtherFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/OtherFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/OtherFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >编辑器上传文件保存目录</td>
    <td ><%=SysRootDir&"Upload/EditorFiles"%></td>
    <td ><%=SizeInfo(SizeByte(SysRootDir&"Upload/EditorFiles"))%></td>
    <td ><%=FormatPercent(Percent(SizeByte(SysRootDir&"Upload/EditorFiles")),,-1)%></td>
    <td ><img src="images/bar.gif" width="<%=Percent(SizeByte(SysRootDir&"Upload/EditorFiles"))*260%>" height="8"></td>
  </tr>
  <tr bgcolor="#EBF2F9" onMouseOver = 'this.style.backgroundColor = "#FFFFFF"' onMouseOut = 'this.style.backgroundColor = ""' style="cursor:hand">
    <td height="20" >临时文件</td>
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
    SizeInfo="<font color='red'>未找到</font>"
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