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
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>常量设置</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
</HEAD>
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|114,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
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
       response.write "<script language=javascript> alert('保存失败，请拷贝后重新打开文件再保存。');location.replace('SetConst.asp');</script>"
	else
       response.write "<script language=javascript> alert('站点常量设置修改成功!');location.replace('SetConst.asp');</script>"
	end if
end if
%>

<BODY>
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
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<form name="ConstSetForm" action="SetConst.asp?Result=Modify" method="post">
<textarea name="ConstContent" rows="22" class="ConstSet" style="width:100%;"><% = FileContent %></textarea>
  <tr>
    <td width="10%"><input name="submitSave" type="submit" class="button" id="submitSave" value=" 保存 "></td>
    <td width="90%" align="right"><font color="#FF0000">注意：主窗口里任意单引号"<font color="#0000FF">'</font>"、"<font color="#0000FF">&lt;%</font>"和"<font color="#0000FF">%&gt;</font>"不能去掉，建议只修改字符，不要增加、删除及使用回车键!</font></td>
  </tr>
</form>
</table>
</body>
</html>