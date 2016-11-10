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
<TITLE>编辑企业</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<%
call CreateEditor("Content")
%>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|11,")=0 Or Instr(session("AdminPurview"),"|300,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,AboutName,ViewFlag,Content
dim GroupID,GroupIdName,Exclusive,ChildFlag
ID=request.QueryString("ID")
call AboutEdit() 
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>企业信息：添加，修改介绍企业相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AboutEdit.asp?Result=Add" onClick='changeAdminFlag("添加企业信息")'>添加企业信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="AboutList.asp" onClick='changeAdminFlag("企业信息")'>查看企业信息</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editAboutForm" method="post" action="AboutEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">信息名称：</td>
        <td><input name="AboutName" type="text" class="textfield" id="AboutName" style="WIDTH: 240;" value="<%=AboutName%>" maxlength="100">&nbsp;发布：<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag then response.write ("checked")%>>
&nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">标　　记：</td>
        <td><input name="ChildFlag" type="checkbox" value="1" style='HEIGHT: 13px;WIDTH: 13px;' <%if ChildFlag then response.write ("checked")%>>&nbsp;分页</td>
      </tr>
      <tr>
        <td height="20" align="right">查看权限：</td>
        <td><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> 隶属<input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
      </tr>
      <tr>
        <td height="20" rowspan="2" align="right" valign="top">信息内容：<br>
        
        <td style="padding:6px"><textarea name="Content" rows="30" class="textfield" id="Content" style="WIDTH: 86%;"  ><%=Content%></textarea></td>
      </tr>
      <tr>
        <td></td>
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
sub AboutEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("AboutName")))="" then
      response.write ("<script language=javascript> alert('信息名称为必填项目！');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '创建网站管理员
	  sql="select * from NwebCn_About"
      rs.open sql,conn,1,3
      rs.addnew
      rs("AboutName")=trim(Request.Form("AboutName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  if Request.Form("ChildFlag")=1 then
        rs("ChildFlag")=Request.Form("ChildFlag")
	    rs("Sequence")=999
	  else
        rs("ChildFlag")=0
	    rs("Sequence")=99
	  end if
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_About where ID="&ID
      rs.open sql,conn,1,3
      rs("AboutName")=trim(Request.Form("AboutName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  if Request.Form("ChildFlag")=1 then
        rs("ChildFlag")=Request.Form("ChildFlag")
	    rs("Sequence")=100
	  else
        rs("ChildFlag")=0
	  end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑企业信息！');changeAdminFlag('企业信息列表');location.replace('AboutList.asp');</script>"
  else '提取管理员信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_About where ID="& ID
      rs.open sql,conn,1,1
	  AboutName=rs("AboutName")
	  ViewFlag=rs("ViewFlag")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
      Content=rs("Content")
	  ChildFlag=rs("ChildFlag")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>
<% 
sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from NwebCn_MemGroup"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
    if GroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
end sub
%>
