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
<TITLE>下载分类</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|51,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<%
Dim Action
Action=request.QueryString("Action")
Select Case Action
  Case "Add"
	addFolder
  	CallFolderView()
  Case "Del"
    Dim rs,sql,SortPath
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From NwebCn_DownSort where ID="&request.QueryString("id")
    rs.open sql,conn,1,1
	SortPath=rs("SortPath")
	conn.execute("delete from  NwebCn_DownSort  where Instr(SortPath,'"&SortPath&"')>0")
    conn.execute("delete from  NwebCn_Download where Instr(SortPath,'"&SortPath&"')>0")
    response.write ("<script language=javascript> alert('成功删除本类、子类及所有下属信息条目，点击确定查看类别树！');location.replace('DownSort.asp');</script>")
  Case "Save"
	saveFolder ()
  Case "Edit"
	editFolder
  	CallFolderView()	
  Case "Move"
	moveFolderForm ()
  	CallFolderView()
  Case "MoveSave"
	saveMoveFolder ()
  Case Else
	CallFolderView()
End Select
%>
</BODY>
</HTML>
<%
'调用显示节点------------------------------
Function CallFolderView()
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><strong>类别树查看管理：</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="DownSort.asp?Action=Add&ParentID=0">添加一级分类</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DownList.asp" onClick='changeAdminFlag("下载列表")'>查看所有下载</a></td>
  </tr>
  <tr>
    <td height="24" nowrap  bgcolor="#EBF2F9"><% Folder(0) %></td>
  </tr>
</table>
<%
End Function
'列出所有节点------------------------------
Function Folder(id)
  Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_DownSort where ParentID="&id&" order by id"
  rs.open sql,conn,1,1
  if id=0 and rs.recordcount=0 then
    response.write ("暂无分类!")
    response.end
  end if  
  i=1
  response.write("<table border='0' cellspacing='0' cellpadding='0'>")
  while not rs.eof
    ChildCount=conn.execute("select count(*) from NwebCn_DownSort where ParentID="&rs("id"))(0)
    if ChildCount=0 then
	  if i=rs.recordcount then
	    FolderType="SortFileEnd"
	  else
	    FolderType="SortFile"
	  end if
	  FolderName=rs("SortName")
	  onMouseUp=""
    else
	  if i=rs.recordcount then
	 	FolderType="SortEndFolderClose"
		ListType="SortEndListline"
		onMouseUp="EndSortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  else
		FolderType="SortFolderClose"
		ListType="SortListline"
		onMouseUp="SortChange('a"&rs("id")&"','b"&rs("id")&"');"
	  end if
	  FolderName=rs("SortName")
    end if
    response.write("<tr>")
    response.write("<td nowrap id='b"&rs("id")&"' class='"&FolderType&"' onMouseUp="&onMouseUp&"></td><td nowrap>"&FolderName&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")	
    response.write("<font color='#FF0000'>分类：</font><a href='DownSort.asp?Action=Add&ParentID="&rs("id")&"'>添加</a>")
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='DownSort.asp?Action=Edit&ID="&rs("id")&"'>修改</a>")
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='DownSort.asp?Action=Move&ID="&rs("id")&"&ParentID="&rs("Parentid")&"&SortName="&rs("SortName")&"&SortPath="&rs("SortPath")&"'>移</a>")
    response.write("→<a href='#' onclick='SortFromTo.rows[4].cells[0].innerHTML=""→&nbsp;"&rs("SortName")&""";MoveForm.toID.value="&rs("ID")&";MoveForm.toParentID.value="&rs("ParentID")&";MoveForm.toSortPath.value="""&rs("SortPath")&""";'>至</a>")
	response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href=javascript:ConfirmDelSort('DownSort',"&rs("id")&")>删除</a>")
    response.write("&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>下载：</font><a href='DownEdit.asp?Result=Add' onClick='changeAdminFlag(""添加下载"")'>添加</a>")
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='DownList.asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"' onClick='changeAdminFlag(""下载列表"")'>列表</a>")
    response.write("</td></tr>")
    if ChildCount>0 then
%>
      <tr id="a<%= rs("id")%>" style="display:yes"><td class="<%= ListType%>" nowrap></td><td ><% Folder(rs("id")) %></td></tr>
<%
    end if
    rs.movenext
    i=i+1
  wend
  response.write("</table>")
  rs.close
  set rs=nothing
end function
'添加节点---------------------------------
Function addFolder()
  Dim ParentID
  ParentID=request.QueryString("ParentID")
  addFolderForm ParentID
end function
'添加节点表单------------------------------
Function addFolderForm(ParentID)
  Dim ParentPath,SortTextPath,rs,sql
  if ParentID=0 then
    ParentPath="0,"
	SortTextPath=""
  else 
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From NwebCn_DownSort where ID="&ParentID
    rs.open sql,conn,1,1
	ParentPath=rs("SortPath")
  end if
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
<form name="FolderForm" method="post" action="DownSort.asp?Action=Save&From=Add">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>添加类别：通过"发布"可控制每种分类是否在相应语言版网站里显示出来。</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">|&nbsp;根类&nbsp;→&nbsp;<% if ParentID<>0 then TextPath(ParentID)%></td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="190" nowrap>名称：<input name="SortName" type="text" class="textfield" id="SortName" size="22"></td>
        <td width="130" nowrap>发布：<input name="ViewFlag" type="radio" value="1" checked="checked" />是<input name="ViewFlag" type="radio" value="0" />否</td>
        <td width="120" nowrap>父类ID：<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="6" value="<%=ParentID %>"></td>
        <td nowrap>父类数字路径：<input readonly name="ParentPath" type="text" class="textfield" id="ParentPath" size="45" value="<%=ParentPath%>"></td>
	  </tr>
      <tr>
        <td colspan="4" align="center" height="30" valign="bottom" nowrap><input name="submitSave" type="submit" class="button" id="保存" value="  保存  "></td>
	  </tr>
    </table>
	</td>
  </tr>
</form>
</table>
<br>
<%
End Function
'生成节点文字路径--------------------------
Function TextPath(ID)
  Dim rs,sql,SortTextPath
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_DownSort where ID="&ID
  rs.open sql,conn,1,1
  SortTextPath=rs("SortName")&"&nbsp;→&nbsp;"
  if rs("ParentID")<>0 then TextPath rs("ParentID")
  response.write(SortTextPath)
End Function
'保存添加、修改节点-------------------------
Function saveFolder
  if len(trim(request.Form("SortName")))=0 then
      response.write ("<script language=javascript> alert('类别名称为必填项目！');history.back(-1);</script>")
      response.end
  end if
  Dim From,Action,rs,sql,SortTextPath
  From=request.QueryString("From")
  Set rs=server.CreateObject("adodb.recordset")
  if From="Add" then 
    sql="Select * From NwebCn_DownSort"
    rs.open sql,conn,1,3
    rs.addnew
	Action="添加类别"
    rs("SortPath")=request.Form("ParentPath") & rs("ID") &","
  else
    sql="Select * From NwebCn_DownSort where ID="&request.QueryString("ID")
    rs.open sql,conn,1,3
	Action="修改类别"
    rs("SortPath")=request.Form("SortPath")
  end if
  rs("SortName")=request.Form("SortName")
  rs("ViewFlag")=request.Form("ViewFlag")
  rs("ParentID")=request.Form("ParentID")
  rs.update 
  response.write ("<script language=javascript> alert('"&Action&"保存成功，点击确定查看类别树！');location.replace('DownSort.asp');</script>")
End Function 
'修改节点---------------------------------
Function editFolder()
  Dim ID
  ID=request.QueryString("ID")
  editFolderForm ID
end function
'修改节点表单------------------------------
Function editFolderForm(ID)
  Dim SortName,ViewFlag,ParentID,SortPath,rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_DownSort where ID="&ID
  rs.open sql,conn,1,1
  SortName=rs("SortName")
  ViewFlag=rs("ViewFlag")
  ParentID=rs("ParentID")
  SortPath=rs("SortPath")
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
<form name="FolderForm" method="post" action="DownSort.asp?Action=Save&From=Edit&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>修改类别：通过"发布"可控制每种分类是否在网站里显示出来。</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">|&nbsp;根类&nbsp;→&nbsp;<% if ParentID<>0 then TextPath(ParentID)%></td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="190" nowrap>名称：<input name="SortName" type="text" class="textfield" id="SortName" size="22" value="<%=SortName%>"></td>
        <td width="130" nowrap>发布：<input name="ViewFlag" type="radio" value="1" <%if ViewFlag then response.write ("checked=checked")%> />是<input name="ViewFlag" type="radio" value="0" <%if not ViewFlag then response.write ("checked=checked")%>/>否</td>
        <td width="120" nowrap>父类ID：<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="6" value="<%=ParentID %>"></td>
        <td nowrap>本类数字路径：<input readonly name="SortPath" type="text" class="textfield" id="SortPath" size="45" value="<%=SortPath%>"></td>
	  </tr>
      <tr>
        <td colspan="4" align="center" height="30" valign="bottom" nowrap><input name="submitSave" type="submit" class="button" id="保存" value="  保存  "></td>
	  </tr>
    </table>
	</td>
  </tr>
</form>
</table>
<br>
<%
End Function
'转移节点表单------------------------------
Function moveFolderForm()
  Dim ID,ParentID,SortName,SortPath
  ID=request.QueryString("ID")
  ParentID=request.QueryString("ParentID")
  SortName=request.QueryString("SortName")
  SortPath=request.QueryString("SortPath")
%>
<table id="SortFromTo" width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
<form name="MoveForm" method="post" action="DownSort.asp?Action=MoveSave">
  <tr>
    <td height="24" colspan="3" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>类别移动：通过点击分类树中类别对应的"移"可重新选择将要作移动的类别，包括本类、子类及所有下属信息条目将一起被移动。</strong></font></td>
  </tr>
  <tr>
    <td height="24" colspan="3" nowrap bgcolor="#EBF2F9">→&nbsp;<% response.write (SortName) %></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#EBF2F9">移动类ID：<input readonly name="ID" type="text" class="textfield" id="ID" size="14" value="<%=ID%>"></td>
    <td nowrap bgcolor="#EBF2F9">移动类父ID：<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="14" value="<%=ParentID%>"></td>
    <td nowrap bgcolor="#EBF2F9">移动类数字路径：<input readonly name="SortPath" type="text" class="textfield" id="SortPath" size="30" value="<%=SortPath%>"></td>
  </tr>
  <tr>
    <td height="24" colspan="3" nowrap><font color="#FFFFFF"><strong>目标位置：通过点击"至"选择将要放置到的类别。</strong></font></td>
  </tr>
  <tr>
    <td height="24" colspan="3" nowrap bgcolor="#EBF2F9">→&nbsp;请选择…</td>
  </tr>
  <tr>
    <td nowrap bgcolor="#EBF2F9">目标类ID：<input readonly name="toID" type="text" class="textfield" id="toID" size="14" value=""></td>
    <td nowrap bgcolor="#EBF2F9">目标类父ID：<input readonly name="toParentID" type="text" class="textfield" id="toParentID" size="14" value=""></td>
    <td nowrap bgcolor="#EBF2F9">目标类数字路径：<input readonly name="toSortPath" type="text" class="textfield" id="toSortPath" size="30" value=""></td>
  </tr>
  <tr>
    <td height="40" colspan="3" nowrap bgcolor="#EBF2F9" align="center"><input name="submitMove" type="submit" class="button" id="转移" value="  转移  "></td>
  </tr>
</form>
</table>
<br>
<%
End Function
'保存转移节点------------------------------
Function saveMoveFolder()
  Dim rs,sql,fromID,fromParentID,fromSortPath,toID,toParentID,toSortPath,fromParentSortPath
  fromID=request.Form("ID")
  fromParentID=request.Form("ParentID")
  fromSortPath=request.Form("SortPath")
  toID=request.Form("toID")
  toParentID=request.Form("toParentID")
  toSortPath=request.Form("toSortPath")
  if toID="" or toParentID="" or toSortPath="" then
    response.write ("<script language=javascript> alert('没有选择移动的目标位置，请返回选择！');history.back(-1);</script>")
    response.end
  end if
  if fromParentID=0 then
    response.write ("<script language=javascript> alert('一级分类不能被移动，请返回选择！');history.back(-1);</script>")
    response.end
  end if
  if fromSortPath=toSortPath then
    response.write ("<script language=javascript> alert('选择的移动类别和目标位置相同了，请返回重新选择！');history.back(-1);</script>")
    response.end
  end if
  if Instr(toSortPath,fromSortPath)>0 or fromParentID=toID then
    response.write ("<script language=javascript> alert('不能将类别移动到本类或下属类里，请返回重新选择！');history.back(-1);</script>")
    response.end
  end if
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_DownSort where ID="&fromParentID
  rs.open sql,conn,0,1
  fromParentSortPath=rs("SortPath")
  conn.execute("update NwebCn_DownSort set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")'更新类别数字路径
  conn.execute("update NwebCn_DownSort set ParentID='"&toID&"' where ID="&fromID)'更新类别父类ID
  conn.execute("update NwebCn_Download set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")'更新信息数字路径
  response.write ("<script language=javascript> alert('移动类别成功，点击确定查看类别树！');location.replace('DownSort.asp');</script>")
End Function
%>