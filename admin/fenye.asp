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
<TITLE>网站信息设置</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|112,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<body>
<%
dim ID,SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,IcpNumber,SystemSN,syimg
dim MesViewFlag
dim procount,newscount,otherscount,downcount,needcount,messagecount,jobcount
select case request.QueryString("Action")
  case "Save"
    SaveSiteInfo
  case else
    ViewSiteInfo
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
  <form name="editForm" method="post" action="fenye.asp?Action=Save" >
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	   <tr>
        <td height="20" align="right">产品每页条数：</td>
        <td><input name="procount" type="text" class="textfield" id="procount" style="WIDTH: 200;" value="<%=procount%>">&nbsp;*&nbsp;</td>
      </tr>
	   <tr>
        <td height="20" align="right">新闻每页条数：</td>
        <td><input name="newscount" type="text" class="textfield" id="newscount" style="WIDTH: 200;" value="<%=newscount%>">&nbsp;*&nbsp;</td>
      </tr>
	   <tr>
        <td height="20" align="right">下载每页条数：</td>
        <td><input name="downcount" type="text" class="textfield" id="downcount" style="WIDTH: 200;" value="<%=downcount%>">&nbsp;*&nbsp;</td>
      </tr>
	   <tr>
        <td height="20" align="right">需求每页条数：</td>
        <td><input name="needcount" type="text" class="textfield" id="needcount" style="WIDTH: 200;" value="<%=needcount%>">&nbsp;*&nbsp;</td>
      </tr>
	   <tr>
        <td height="20" align="right">留言每页条数：</td>
        <td><input name="messagecount" type="text" class="textfield" id="messagecount" style="WIDTH: 200;" value="<%=messagecount%>">&nbsp;*&nbsp;</td>
      </tr>
	   <tr>
        <td height="20" align="right">其他每页条数：</td>
        <td><input name="otherscount" type="text" class="textfield" id="otherscount" style="WIDTH: 200;" value="<%=otherscount%>">&nbsp;*&nbsp;</td>
      </tr>
	  <tr>
        <td height="20" align="right">招聘每页条数：</td>
        <td><input name="jobcount" type="text" class="textfield" id="jobcount" style="WIDTH: 200;" value="<%=jobcount%>">&nbsp;*&nbsp;</td>
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
function SaveSiteInfo()

  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from NwebCn_Site"
  rs.open sql,conn,1,3
 
  rs("procount")=trim(Request.Form("procount"))
  rs("newscount")=trim(Request.Form("newscount"))
  rs("otherscount")=trim(Request.Form("otherscount"))
  rs("downcount")=trim(Request.Form("downcount"))
  rs("needcount")=trim(Request.Form("needcount"))
  rs("messagecount")=trim(Request.Form("messagecount"))
  rs("jobcount")=trim(Request.Form("jobcount"))
rs.update
  rs.close
  set rs=nothing 
  response.write "<script language=javascript> alert('成功编辑网站信息！');changeAdminFlag('网站信息设置');location.replace('fenye.asp');</script>"
end function

function ViewSiteInfo()
  dim rs,sql 
  set rs = server.createobject("adodb.recordset")
  sql="select top 1 * from NwebCn_Site"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write "读取数据库记录出错！"
    response.end
  else

	procount=rs("procount")
	newscount=rs("newscount")
	otherscount=rs("otherscount")
	downcount=rs("downcount")
	needcount=rs("needcount")
	messagecount=rs("messagecount")
	jobcount=rs("jobcount")
    rs.close
    set rs=nothing 
  end if
end function
%>
