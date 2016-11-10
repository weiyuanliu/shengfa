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
dim ID,SiteName,ViewFlag,LinkType,SiteFace,SiteUrl,Remark
ID=request.QueryString("ID")
call FriendSiteEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>服务品牌：添加，修改友情链接相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="FriendSiteEdit.asp?Result=Add" onClick='changeAdminFlag("添加服务品牌")'>添加友情连接</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="FriendSiteList.asp" onClick='changeAdminFlag("服务品牌列表")'>查看友情连接</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="FriendSiteEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">鼠标移动介绍：</td>
        <td><input name="SiteName" type="text" class="textfield" id="SiteName" style="WIDTH: 240;" value="<%=SiteName%>">&nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">发　　布：</td>
        <td><input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag then response.write ("checked")%>></td>
      </tr>
	   <tr>
        <td height="20" align="right">排序：</td>
        <td><input name="px" type="text" class="textfield" id="px" style="WIDTH: 100;" value="<%=px%>">          &nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">链接类型：</td>
        <td><input name="LinkType" type="radio" value="1" <%if LinkType then response.write ("checked=checked")%>/>图片&nbsp;<input name="LinkType" type="radio" value="0" <%if not LinkType then response.write ("checked=checked")%>/>文字</td>
      </tr>
      <tr>
        <td height="20" align="right">前台显示：</td>
        <td><input name="SiteFace" type="text" class="textfield" id="SiteFace" style="WIDTH: 240;" value="<%=SiteFace%>">
        &nbsp;*&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SiteFace',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a>&nbsp;&nbsp;图片196×67&nbsp;&nbsp;文字≤8个汉字</td>
      </tr>
      <tr>
        <td height="20" align="right">链接网址：</td>
        <td><input name="SiteUrl" type="text" class="textfield" id="SiteUrl" style="WIDTH: 490;" value="<%=SiteUrl%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">备注说明：
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 490;"><%=Remark%></textarea></td>
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
sub FriendSiteEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("SiteName")))<2 then
      response.write ("<script language=javascript> alert('网站名称为必填项目且不少于2个字符！');history.back(-1);</script>")
      response.end
    end if
    if trim(request.Form("SiteFace"))="" then
      response.write ("<script language=javascript> alert('前台显示为必填项目！');history.back(-1);</script>")
      response.end
    end if
    if request.Form("LinkType")=0 then
      if StrLen(trim(request.Form("SiteFace")))>16 then
      response.write ("<script language=javascript> alert('您选择的""文字""链接，因此前台显示不得超过8个汉字！');history.back(-1);</script>")
      response.end
      end if
    end if
    if len(trim(request.Form("SiteUrl")))<6 then
      response.write ("<script language=javascript> alert('链接网址为必填项目且不少于6个字符！');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '创建网站管理员
	  sql="select * from NwebCn_FriendSite"
      rs.open sql,conn,1,3
      rs.addnew
      rs("SiteName")=trim(Request.Form("SiteName"))
	  if isnumeric(trim(Request.Form("Px"))) then
	  rs("Px")=trim(Request.Form("Px"))
	  else
	  rs("Px")=0
	  end if
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
      rs("SiteFace")=trim(Request.Form("SiteFace"))
      rs("SiteUrl")=trim(Request.Form("SiteUrl"))
	  if Request.Form("LinkType")=1 then
        rs("LinkType")=Request.Form("LinkType")
	  else
        rs("LinkType")=0
	  end if	  
	  rs("Remark")=trim(Request.Form("Remark"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_FriendSite where ID="&ID
      rs.open sql,conn,1,3
      rs("SiteName")=trim(Request.Form("SiteName"))
	   if isnumeric(trim(Request.Form("Px"))) then
	  rs("Px")=trim(Request.Form("Px"))
	  else
	  rs("Px")=0
	  end if
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
      rs("SiteFace")=trim(Request.Form("SiteFace"))
      rs("SiteUrl")=trim(Request.Form("SiteUrl"))
	  if Request.Form("LinkType")=1 then
        rs("LinkType")=Request.Form("LinkType")
	  else
        rs("LinkType")=0
	  end if	  
	  rs("Remark")=trim(Request.Form("Remark"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑友情链接！');changeAdminFlag('友情链接列表');location.replace('FriendSiteList.asp');</script>"
  else '提取信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_FriendSite where ID="& ID
      rs.open sql,conn,1,1
	  SiteName=rs("SiteName")
	  ViewFlag=rs("ViewFlag")
	  LinkType=rs("LinkType")
	  SiteFace=rs("SiteFace")
      SiteUrl=rs("SiteUrl")
      Remark=rs("Remark")
	  Px=rs("Px")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>