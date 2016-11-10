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
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>修改密码</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|111,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
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
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>系统管理：添加，修改站点的相关信息</strong></font></td>
  </tr>
 <!-- <tr>
    <td height="24" align="center" nowrap bgcolor="#EBF2F9">
	<a href="PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("修改密码")'>修改密码</a>	<font color="#0000FF">&nbsp;|&nbsp;</font>	<a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>网站信息设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" target="mainFrame" onClick='changeAdminFlag("栏目导航设置")'>栏目导航设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetConst.asp" target="mainFrame" onClick='changeAdminFlag("常量设置")'>常量设置</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DataManage.asp" target="mainFrame" onClick='changeAdminFlag("数据库操作")'>数据库操作</a>
<font color="#0000FF">&nbsp;|&nbsp;</font><a href="ADsEdit.asp?Result=Add" target="mainFrame" onClick='changeAdminFlag("弹窗广告列表")'>弹窗广告</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SpaceStat.asp" target="mainFrame" onClick='changeAdminFlag("空间统计")'>空间统计</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="../Count/InfoList.asp" target="mainFrame" onClick='changeAdminFlag("访问统计")'>访问统计</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="FriendSiteList.asp" target="mainFrame" onClick='changeAdminFlag("友情链接")'>友情链接</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="HackSql.asp" target="mainFrame" onClick='changeAdminFlag("阻止SQL注入记录")'>阻止SQL注入记录</a>    </td>
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
        <td height="20" align="right">登&nbsp;录&nbsp;名：</td>
        <td><input name="AdminName" type="text" class="textfield" id="AdminName" style="WIDTH: 120;" value="<%=session("AdminName")%>" maxlength="16" readonly>&nbsp;3-10位字符，不可修改</td>
      </tr>
      <tr>
        <td height="20" align="right">新&nbsp;密&nbsp;码：</td>
        <td><input name="NewPassword" type="password" class="textfield" id="NewPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*&nbsp;注意字母大小写</td>
      </tr>
      <tr>
        <td height="20" align="right">确认密码：</td>
        <td><input name="vNewPassword" type="password" class="textfield" id="vNewPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*</td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 60;" ></td>
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
    response.write "读取数据库记录出错！"
    response.end
  else
	if len(trim(Request.Form("NewPassword")))<6 or len(trim(Request.Form("NewPassword")))>20  then
      response.write "<script language=javascript> alert('管理员密码必填，且字符数为6-20位！');history.back(-1);</script>"
      response.end
    end if
	if Request.Form("NewPassword")<>Request.Form("vNewPassword") then 
      response.write "<script language=javascript> alert('两次输入的密码不一样！');history.back(-1);</script>"
      response.end
	end if
	rs("Password")=Md5(Request.Form("NewPassword"))  
    rs.update
    rs.close
    set rs=nothing 
  end if
  response.write "你的管理密码已成功修改，请牢记[&nbsp;<font color='red'>"&trim(Request.Form("NewPassword"))&"</font>&nbsp;]！"
  response.end
end function
%>
