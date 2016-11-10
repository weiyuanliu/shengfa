<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七日科技企业网站管理系统（LISuo）　　　　　　　  ┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
' 　版权所有　qisehu.com
'
'　　程序制作　七色狐网络有限公司
'　　　　　　　Add:四川省成都市二环路西三段181号13楼20/21号
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七色狐网络有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑友情链接</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|119,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result,px
Result=request.QueryString("Result")
dim ID,ADS_Name,ADS_Link,AddTime
ID=request.QueryString("ID")
call ADVEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>广告：添加，修改广告相关的内容</strong></font></td>
  </tr>
  <tr>
        <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="advset.asp?Result=Add" onClick='changeAdminFlag("添加广告")'>添加广告</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="advlist.asp" onClick='changeAdminFlag("广告列表")'>查看广告</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="advset.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">广告名称：</td>
        <td><input name="ADS_Name" type="text" class="textfield" id="ADS_Name" style="WIDTH: 240;" value="<%=ADS_Name%>">&nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">链接网址：</td>
        <td><input name="ADS_Link" type="text" class="textfield" id="ADS_Link" style="WIDTH: 490;" value="<%=ADS_Link%>">&nbsp;*</td>
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
function ADVEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("ADS_Name")))<2 then
      response.write ("<script language=javascript> alert('广告名称为必填项目且不少于2个字符！');history.back(-1);</script>")
      response.end
    end if
    if len(trim(request.Form("ADS_Link")))<2 then
      response.write ("<script language=javascript> alert('广告地址为必填项目且不少于10个字符！');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then 
	  sql="select * from NwebCn_Ads_effect"
      rs.open sql,conn,1,3
      rs.addnew
      rs("ADS_Name")=trim(Request.Form("ADS_Name"))
      rs("ADS_Link")=trim(Request.Form("ADS_Link"))
      rs("AddTime")=now
	end if
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Ads_effect where ID="&ID
      rs.open sql,conn,1,3
      rs("ADS_Name")=trim(Request.Form("ADS_Name"))
      rs("ADS_Link")=trim(Request.Form("ADS_Link"))
	end if
	  rs.update
	  rs.close
      set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑广告！');changeAdminFlag('广告列表');location.replace('advlist.asp');</script>"
  else '提取信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Ads_effect where ID="& ID
      rs.open sql,conn,1,1
	  ADS_Name=rs("ADS_Name")
	  ADS_Link=rs("ADS_Link")
	  rs.close
      set rs=nothing 
	end if
  end if
end function
%>