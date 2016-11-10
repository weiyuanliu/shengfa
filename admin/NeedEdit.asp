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
<TITLE>编辑需求</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|43,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,NeedName,ViewFlag,SortName,SortID,SortPath
dim UrgentFlag,GroupID,GroupIdName,Exclusive,EndDate,Content
ID=request.QueryString("ID")
call NeedEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>需求检索及分类查看：添加，修改，删除需求信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="NeedEdit.asp?Result=Add" onClick='changeAdminFlag("添加需求信息")'>添加需求信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NeedList.asp" onClick='changeAdminFlag("需求列表")'>查看所有需求信息</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="NeedEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNeed idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">需求名称：</td>
        <td><input name="NeedName" type="text" class="textfield" id="NeedName" style="WIDTH: 240;" value="<%=NeedName%>" maxlength="100">&nbsp;显示：<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag or Result="Add" then response.write ("checked")%>>&nbsp;*&nbsp;不少于3个字符</td>
      </tr>
      <tr>
        <td height="20" align="right">所属类别：</td>
        <td><input name="SortName" type="text" class="textfield" id="SortNameSi" value="<%=SortName%>" style="WIDTH: 240;background-color:#EBF2F9;" readonly>&nbsp;<a href="javaScript:OpenScript('SelectSort.asp?Result=Need',500,500,'')"><img src="Images/Select.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">类别数字：</td>
        <td><input name="SortID" type="text" class="textfield" id="SortID" style="WIDTH: 40;background-color:#EBF2F9;" value="<%=SortID%>" readonly><input name="SortPath" type="text" class="textfield" id="SortPath" style="WIDTH: 200;background-color:#EBF2F9;" value="<%=SortPath%>" readonly>&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">查看权限：</td>
        <td><select name="GroupID" class="textfield">
            <% call SelectGroup() %>
          </select>
            <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>>
          隶属
          <input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">
          专属（隶属：权限值≥可查看，专属：权限值＝可查看）</td>
      </tr>
      <tr>
        <td height="20" align="right">状　　态：</td>
        <td><input name="UrgentFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if UrgentFlag then response.write ("checked")%>>&nbsp;紧急</td>
      </tr>
      <tr>
        <td height="20" align="right">结束日期：</td>
        <td><input name="EndDate" type="text" class="textfield" id="EndDate" style="WIDTH: 240;" value="<% if EndDate="" then response.write (DateAdd("m",3,now())) else response.write (EndDate) end if%>" maxlength="14">&nbsp;默认为3个月，可手动输入日期格式</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">简体介绍：<br>
		  <img title="点击进入可视化查看、编辑环境..." src="Images/Edit.gif" width="51" height="20" style="cursor:hand" onClick="OpenDialog('../Editor/EditorDialog.html?lnk=Content&file=Editor_1.html',800,520);">
        <td><textarea name="Content" rows="8" class="textfield" id="Content" style="WIDTH: 86%;" readonly><%=Content%></textarea></td>
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
sub NeedEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑下载信息
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("NeedName")))<3 then
      response.write ("<script language=javascript> alert('需求名称为必填项目！');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '创建需求信息
	  sql="select * from NwebCn_Need"
      rs.open sql,conn,1,3
      rs.addnew
      rs("NeedName")=trim(Request.Form("NeedName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")="" and Request.Form("SortPath")="" then
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back(-1);</script>")
        response.end
	  else
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  end if
	  if Request.Form("UrgentFlag")=1 then
        rs("UrgentFlag")=Request.Form("UrgentFlag")
	  else
        rs("UrgentFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("EndDate")=CDate(trim(Request.Form("EndDate")))
	  rs("Content")=Request.Form("Content")
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改下载信息
      sql="select * from NwebCn_Need where ID="&ID
      rs.open sql,conn,1,3
      rs("NeedName")=trim(Request.Form("NeedName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('请选择所属分类！');history.back(-1);</script>")
        response.end
	  end if
	  if Request.Form("UrgentFlag")=1 then
        rs("UrgentFlag")=Request.Form("UrgentFlag")
	  else
        rs("UrgentFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("EndDate")=CDate(trim(Request.Form("EndDate")))
	  rs("Content")=Request.Form("Content")
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑需求信息！');changeAdminFlag('需求列表');location.replace('NeedList.asp');</script>"
  else '提取需求信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Need where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
      end if
	  NeedName=rs("NeedName")
	  ViewFlag=rs("ViewFlag")
	  SortName=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  SortPath=rs("SortPath")
	  UrgentFlag=rs("UrgentFlag")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
	  EndDate=rs("EndDate")
      Content=rs("Content")
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
<%
'生成所属类别--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_NeedSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>