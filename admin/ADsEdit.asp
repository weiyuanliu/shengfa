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
<TITLE>编辑广告</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,ADsName,ViewFlag,Content
dim ADsWidth,ADsHeight
ID=request.QueryString("ID")
call ADsEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>弹窗广告：添加，修改弹窗广告相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="ADsEdit.asp?Result=Add" onClick='changeAdminFlag("添加弹窗广告")'>添加弹窗广告</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="ADsList.asp" onClick='changeAdminFlag("弹窗广告列表")'>查看弹窗广告</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="ADsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">广告标题：</td>
        <td><input name="ADsName" type="text" class="textfield" id="ADsName" style="WIDTH: 240;" value="<%=ADsName%>" maxlength="100">&nbsp;发布：<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag then response.write ("checked")%>>
&nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">弹窗尺寸：</td>
        <td><input name="ADsWidth" type="text" class="textfield" id="ADsWidth" style="WIDTH: 60;" value="<%=ADsWidth%>" maxlength="4" onKeyDown="if(event.keyCode==13)event.returnValue=false" onChange="if(/\D/.test(this.value)){alert('宽度和高度只能输入整数！');this.value='150';}">&nbsp;宽×高&nbsp;<input name="ADsHeight" type="text" class="textfield" id="ADsHeight" style="WIDTH: 60;" value="<%=ADsHeight%>" maxlength="4" onKeyDown="if(event.keyCode==13)event.returnValue=false" onChange="if(/\D/.test(this.value)){alert('宽度和高度只能输入整数！');this.value='100';}">&nbsp;*&nbsp;至少150×100像素</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">简体内容：<br>
		  <img title="点击进入可视化查看、编辑环境..." src="Images/Edit.gif" width="51" height="20" style="cursor:hand" onClick="OpenDialog('../Editor/EditorDialog.html?lnk=Content&file=Editor_1.html',800,520);">
        <td><textarea name="Content" rows="12" class="textfield" id="Content" style="WIDTH: 86%;" readonly><%=Content%></textarea></td>
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
sub ADsEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("ADsName")))<3 then
      response.write ("<script language=javascript> alert('广告标题为必填项目！');history.back(-1);</script>")
      response.end
    end if
	if trim(request.Form("ADsWidth"))="" or trim(request.Form("ADsHeight"))="" then
      response.write ("<script language=javascript> alert('弹窗广告规格必须为150×100像素以上！');history.back(-1);</script>")
      response.end
	end if
	if trim(request.Form("ADsWidth"))<150 or trim(request.Form("ADsHeight"))<100 then
      response.write ("<script language=javascript> alert('弹窗广告规格必须为150×100像素以上！');history.back(-1);</script>")
      response.end
	end if
    if Result="Add" then '创建网站管理员
	  sql="select * from NwebCn_ADs"
      rs.open sql,conn,1,3
      rs.addnew
      rs("ADsName")=trim(Request.Form("ADsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
	  rs("ADsWidth")=trim(Request.Form("ADsWidth"))
	  rs("ADsHeight")=trim(Request.Form("ADsHeight"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_ADs where ID="&ID
      rs.open sql,conn,1,3
      rs("ADsName")=trim(Request.Form("ADsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
	  rs("ADsWidth")=trim(Request.Form("ADsWidth"))
	  rs("ADsHeight")=trim(Request.Form("ADsHeight"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑弹窗广告！');changeAdminFlag('弹窗广告列表');location.replace('ADsList.asp');</script>"
  else '提取管理员信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_ADs where ID="& ID
      rs.open sql,conn,1,1
	  ADsName=rs("ADsName")
	  ViewFlag=rs("ViewFlag")
	  ADsWidth=rs("ADsWidth")
	  ADsHeight=rs("ADsHeight")
      Content=rs("Content")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>