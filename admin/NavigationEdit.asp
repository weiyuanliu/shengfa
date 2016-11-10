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
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑导航</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|113,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,NavName,ViewFlag,NavUrl,Remark
ID=request.QueryString("ID")
call NavEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>导航栏目：添加，修改导航栏目相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="NavigationEdit.asp?Result=Add" onClick='changeAdminFlag("添加导航栏目")'>添加导航栏目</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" onClick='changeAdminFlag("导航栏目列表")'>查看导航栏目</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="NavigationEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">导航名称：</td>
        <td><input name="NavName" type="text" class="textfield" id="NavName" style="WIDTH: 240;" value="<%=NavName%>" maxlength="100">&nbsp;发布：<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag  or Result="Add" then response.write ("checked")%>>&nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">链接网址：</td>
        <td><input name="NavUrl" type="text" class="textfield" id="NavUrl" style="WIDTH: 480;" value="<%=NavUrl%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">备注说明：
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 480;"><%=Remark%></textarea></td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub NavEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("NavName")))<3 then
      response.write ("<script language=javascript> alert('导航名称为必填项目！');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '创建网站管理员
	  sql="select * from NwebCn_Navigation"
      rs.open sql,conn,1,3
      rs.addnew
      rs("NavName")=trim(Request.Form("NavName"))
      rs("NavUrl")=trim(Request.Form("NavUrl"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	  rs("Sequence")=99
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Navigation where ID="&ID
      rs.open sql,conn,1,3
      rs("NavName")=trim(Request.Form("NavName"))
      rs("NavUrl")=trim(Request.Form("NavUrl"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Remark")=trim(Request.Form("Remark"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑导航栏目！');changeAdminFlag('导航栏目列表');location.replace('NavigationList.asp');</script>"
  else '提取管理员信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Navigation where ID="& ID
      rs.open sql,conn,1,1
	  NavName=rs("NavName")
	  ViewFlag=rs("ViewFlag")
      Remark=rs("Remark")
      NavUrl=rs("NavUrl")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>