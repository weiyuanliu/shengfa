<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'����������������������������������������������������������������
'����������������������������������������������������������������
'�������������������տƼ���ҵ��վ����ϵͳ��LISuo����������������  ��
'����������������������������������������������������������������
' ����Ȩ���С�qisehu.com
'   �������ã����޼������
'�����������������տƼ����޹�˾
'��������������Add:�Ĵ�ʡ�ɶ��ж���·������181��13¥20/21��
'����������������������������������������������������������������
'����������������������������������������������������������������
%>
<% Option Explicit %>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|21,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>


<%
Dim Action,px,tb,tbs,sorts,sortList,sortEdit

Action=request.QueryString("Action")
TbS=Request.QueryString("TbS")
Tb=Request.QueryString("Tb")
select Case Tb
	Case "NwebCn_News"
	sorts="����"
	SortList="NewsList.asp"
	sortEdit="NewsEdit.asp"
	Case "NwebCn_Products"
	sorts="��Ʒ"
	SortList="ProductList.asp"
	sortEdit="ProductEdit.asp"
	case "NwebCn_Need"
	sorts="����"
	SortList="NeedList.asp"
	sortEdit="NeedEdit.asp"
	case "NwebCn_Download"
	sorts="����"
	SortList="DownList.asp"
	sortEdit="DownEdit.asp"
	case "NwebCn_Others"
	sorts="����"
	SortList="othersList.asp"
	sortEdit="othersedit.asp"
end select	
	%>
	
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE><%=sorts%>����</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<BODY>
<%
Function DelFile(Files)
dim fs,file
Set fs = Server.CreateObject("Scripting.FileSystemObject")
File = Server.MapPath(Files)
on Error Resume Next
fs.DeleteFile File, True 'ǿ��ɾ��ֻ���ļ�
If Err.Number = 53 Then
Response.Write File & "�ļ������ڣ�"
Response.End
Elseif Err.Number = 70 Then
Response.Write File & "�ļ�����Ϊ����״̬��"
Response.End
Elseif Err.Number <> 0 Then
Response.Write "δ֪���󣬴�����룺" & Err.Number
Response.End
Else
Response.Write "�ɹ�ɾ���ļ���" & File
End If
 
End Function

Function DelPic(Dates,Sortpath)

dim sqls,rss
	set rss=server.CreateObject("Adodb.recordset")
	Sqls="select smallpic,bigpic from "&Dates&" where Instr(SortPath,'"&SortPath&"')>0"
	 
	rss.open sqls,conn,1,3            
	if rss.bof and rss.eof then
	else
	while not rss.eof
	if rss("smallpic")=rss("bigpic") then
	 DelFile(Rss("smallpic"))
	 else
	 
	 DelFile(Rss("bigpic"))
	end if
	rss.movenext
	wend
	end if
	rss.close
	set rss=nothing
End Function
 
Select Case Action
  Case "Add"
	addFolder
  	CallFolderView()
  Case "Del"
    Dim rs,sql,SortPath
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From "& Tbs &" where ID="&request.QueryString("id")
    rs.open sql,conn,1,1
	SortPath=rs("SortPath")
	conn.execute("delete from  "& Tbs &" where Instr(SortPath,'"&SortPath&"')>0")
	
	
	DelPic Tb, SortPath  'ɾ��ͼƬ
    conn.execute("delete from "& Tb &" where Instr(SortPath,'"&SortPath&"')>0")
    response.write ("<script language=javascript> alert('�ɹ�ɾ�����ࡢ���༰����������Ϣ��Ŀ�����ȷ���鿴�������');location.replace('Sort.asp?TbS="&TbS&"&Tb="&Tb&"');</script>")
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
'������ʾ�ڵ�------------------------------
Function CallFolderView()
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><strong>������鿴������</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="Sort.asp?Action=Add&ParentID=0&TbS=<%=TbS%>&Tb=<%=Tb%>">����һ������</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="<%=sortlist%>" onClick='changeAdminFlag("<%=sorts%>�б�")'>�鿴����<%=sorts%></a></td>
  </tr>
  <tr>
    <td height="24" nowrap  bgcolor="#EBF2F9"><% Folder(0) %></td>
  </tr>
</table>
<%
End Function
'�г����нڵ�------------------------------
Function Folder(id)
  Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType,px
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "& Tbs &" where ParentID="&id&" order by id"
  rs.open sql,conn,1,1
  if id=0 and rs.recordcount=0 then
    response.write ("���޷���!")
    response.end
  end if  
  i=1
  response.write("<table border='0' cellspacing='0' cellpadding='0'>")
  while not rs.eof
    ChildCount=conn.execute("select count(*) from "& Tbs &" where ParentID="&rs("id"))(0)
    if ChildCount=0 then
	  if i=rs.recordcount then
	    FolderType="SortFileEnd"
	  else
	    FolderType="SortFile"
	  end if
	  FolderName=rs("SortName")
	  onMouseUp=""
	  px=Rs("px")
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
	  px=rs("px")
    end if
    response.write("<tr>")
    response.write("<td nowrap id='b"&rs("id")&"' class='"&FolderType&"' onMouseUp="&onMouseUp&"></td><td nowrap>"&FolderName&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&vbcrlf)	
    response.write("<font color='#FF0000'>���ࣺ</font><a href='Sort.asp?Action=Add&ParentID="&rs("id")&"&TbS="&TbS&"&Tb="&Tb&"'>����</a>"&vbcrlf)
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='Sort.asp?Action=Edit&ID="&rs("id")&"&TbS="&TbS&"&Tb="&Tb&"'>�޸�</a>"&vbcrlf)
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='Sort.asp?Action=Move&ID="&rs("id")&"&TbS="&TbS&"&Tb="&Tb&"&ParentID="&rs("Parentid")&"&SortName="&server.URLEncode(rs("SortName"))&"&SortPath="&rs("SortPath")&"'>��</a>"&vbcrlf)
    response.write("��<a href='#' onclick='SortFromTo.rows[4].cells[0].innerHTML=""��&nbsp;"&server.URLEncode(rs("SortName"))&""";MoveForm.toID.value="&rs("ID")&";MoveForm.toParentID.value="&rs("ParentID")&";MoveForm.toSortPath.value="""&rs("SortPath")&""";'>��</a>"&vbcrlf)
	response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href=javascript:ConfirmDelSort('sort',"&rs("id")&",'"&tbs&"','"&tb&"')>ɾ��</a>"&vbcrlf)
    response.write("&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>"&sorts&"��</font><a href='"&sortEdit&"?Result=Add' onClick='changeAdminFlag("""&sorts&"����"")'>����</a>"&vbcrlf)
    response.write("<font color='#367BDA'>&nbsp;|&nbsp;</font><a href='"&SortList&"?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&"' onClick='changeAdminFlag("""&sorts&"�б�"")'>�б�</a><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;����"&px&vbcrlf)
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
'���ӽڵ�---------------------------------
Function addFolder()
  Dim ParentID
  ParentID=request.QueryString("ParentID")
  addFolderForm ParentID
end function
'���ӽڵ����------------------------------
Function addFolderForm(ParentID)
  Dim ParentPath,SortTextPath,rs,sql
  if ParentID=0 then
    ParentPath="0,"
	SortTextPath=""
  else 
    Set rs=server.CreateObject("adodb.recordset")
    sql="Select * From "& Tbs &" where ID="&ParentID
    rs.open sql,conn,1,1
	ParentPath=rs("SortPath")
  end if
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
<form name="FolderForm" method="post" action="Sort.asp?Action=Save&From=Add&TbS=<%=TbS%>&Tb=<%=Tb%>">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>�������ͨ��"����"�ɿ���ÿ�ַ����Ƿ�����վ����ʾ������</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">|&nbsp;����&nbsp;��&nbsp;<% if ParentID<>0 then TextPath(ParentID)%></td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="190" nowrap>���ƣ�<input name="SortName" type="text" class="textfield" id="SortName" size="22"></td>
        <td width="130" nowrap>��ʾ��<input name="ViewFlag" type="radio" value="1" checked="checked" />��<input name="ViewFlag" type="radio" value="0" />��</td>
        <td width="100" nowrap>����ID��<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="6" value="<%=ParentID %>"></td>
        <td nowrap>��������·����<input readonly name="ParentPath" type="text" class="textfield" id="ParentPath" size="32" value="<%=ParentPath%>"></td>
	   <td nowrap>����<input  name="px" type="text" class="textfield" id="px" size="5" value="<%=px%>"></td>
	  </tr>
	  
      <tr>
        <td colspan="4" align="center" height="30" valign="bottom" nowrap><input name="submitSave" type="submit" class="button" id="����" value="  ����  "></td>
	  </tr>
    </table>
	</td>
  </tr>
</form>
</table>
<br>
<%
End Function
'���ɽڵ�����·��--------------------------
Function TextPath(ID)
  Dim rs,sql,SortTextPath
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "& Tbs &" where ID="&ID
  rs.open sql,conn,1,1
  SortTextPath=rs("SortName")&"&nbsp;��&nbsp;"
  if rs("ParentID")<>0 then TextPath rs("ParentID")
  response.write(SortTextPath)
End Function
'�������ӡ��޸Ľڵ�-------------------------
Function saveFolder
  if len(trim(request.Form("SortName")))=0 then
      response.write ("<script language=javascript> alert('�������Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
  end if
  Dim From,Action,rs,sql,SortTextPath
  From=request.QueryString("From")
  Set rs=server.CreateObject("adodb.recordset")
  if From="Add" then 
    sql="Select * From "& Tbs &""
    rs.open sql,conn,1,3
    rs.addnew
	Action="�������"
    rs("SortPath")=request.Form("ParentPath") & rs("ID") &","
  else
    sql="Select * From "& Tbs &" where ID="&request.QueryString("ID")
    rs.open sql,conn,1,3
	Action="�޸����"
    rs("SortPath")=request.Form("SortPath")
  end if
  rs("SortName")=request.Form("SortName")
  rs("ViewFlag")=request.Form("ViewFlag")
  rs("ParentID")=request.Form("ParentID")
  if isnumeric(request.Form("Px")) then 
     rs("Px")=request.Form("Px")
  else
    rs("px")=0
  end if
    rs.update 
  response.write ("<script language=javascript> alert('"&Action&"����ɹ������ȷ���鿴�������');location.replace('Sort.asp?TbS="&TbS&"&Tb="&Tb&"');</script>")
End Function 
'�޸Ľڵ�---------------------------------
Function editFolder()
  Dim ID
  ID=request.QueryString("ID")
  editFolderForm ID
end function
'�޸Ľڵ����------------------------------
Function editFolderForm(ID)
  Dim SortName,ViewFlag,ParentID,SortPath,rs,sql,px
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "& Tbs &" where ID="&ID
  rs.open sql,conn,1,1
  SortName=rs("SortName")
  ViewFlag=rs("ViewFlag")
  ParentID=rs("ParentID")
  SortPath=rs("SortPath")
  px=rs("px")
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
<form name="FolderForm" method="post" action="Sort.asp?Action=Save&From=Edit&ID=<%=ID%>&TbS=<%=TbS%>&Tb=<%=Tb%>">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><strong>�޸����ͨ��"����"�ɿ���ÿ�ַ����Ƿ�����վ����ʾ������</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">|&nbsp;����&nbsp;��&nbsp;<% if ParentID<>0 then TextPath(ParentID)%></td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="190" nowrap>���ƣ�<input name="SortName" type="text" class="textfield" id="SortName" size="22" value="<%=SortName%>"></td>
        <td width="130" nowrap>������<input name="ViewFlag" type="radio" value="1" <%if ViewFlag then response.write ("checked=checked")%> />��<input name="ViewFlag" type="radio" value="0" <%if not ViewFlag then response.write ("checked=checked")%>/>��</td>
        <td width="100" nowrap>����ID��<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="6" value="<%=ParentID %>"></td>
        <td nowrap>��������·����<input readonly name="SortPath" type="text" class="textfield" id="SortPath" size="32" value="<%=SortPath%>"></td>
		<td nowrap>����<input  name="px" type="text" class="textfield" id="px" size="5" value="<%=px%>"></td>
	  </tr>
      <tr>
        <td colspan="4" align="center" height="30" valign="bottom" nowrap><input name="submitSave" type="submit" class="button" id="����" value="  ����  "></td>
	  </tr>
    </table>
	</td>
  </tr>
</form>
</table>
<br>
<%
End Function
'ת�ƽڵ����------------------------------
Function moveFolderForm()
  Dim ID,ParentID,SortName,SortPath
  ID=request.QueryString("ID")
  ParentID=request.QueryString("ParentID")
  SortName=request.QueryString("SortName")
  SortPath=request.QueryString("SortPath")
%>
<table id="SortFromTo" width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
<form name="MoveForm" method="post" action="Sort.asp?Action=MoveSave&TbS=<%=TbS%>&Tb=<%=Tb%>">
  <tr>
    <td height="24" colspan="3" nowrap><font color="#FFFFFF"><strong>����ƶ���ͨ�����������������Ӧ��"��"������ѡ��Ҫ���ƶ�����𣬰������ࡢ���༰����������Ϣ��Ŀ��һ���ƶ���</strong></font></td>
  </tr>
  <tr>
    <td height="24" colspan="3" nowrap bgcolor="#EBF2F9">��&nbsp;<% response.write (SortName) %></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#EBF2F9">�ƶ���ID��<input readonly name="ID" type="text" class="textfield" id="ID" size="14" value="<%=ID%>"></td>
    <td nowrap bgcolor="#EBF2F9">�ƶ��ุID��<input readonly name="ParentID" type="text" class="textfield" id="ParentID" size="14" value="<%=ParentID%>"></td>
    <td nowrap bgcolor="#EBF2F9">�ƶ�������·����<input readonly name="SortPath" type="text" class="textfield" id="SortPath" size="30" value="<%=SortPath%>"></td>
  </tr>
  <tr>
    <td height="24" colspan="3" nowrap><font color="#FFFFFF"><strong>Ŀ��λ�ã�ͨ�����"��"ѡ��Ҫ���õ������</strong></font></td>
  </tr>
  <tr>
    <td height="24" colspan="3" nowrap bgcolor="#EBF2F9">��&nbsp;��ѡ��</td>
  </tr>
  <tr>
    <td nowrap bgcolor="#EBF2F9">Ŀ����ID��<input readonly name="toID" type="text" class="textfield" id="toID" size="14" value=""></td>
    <td nowrap bgcolor="#EBF2F9">Ŀ���ุID��<input readonly name="toParentID" type="text" class="textfield" id="toParentID" size="14" value=""></td>
    <td nowrap bgcolor="#EBF2F9">Ŀ��������·����<input readonly name="toSortPath" type="text" class="textfield" id="toSortPath" size="30" value=""></td>
  </tr>
  <tr>
    <td height="40" colspan="3" nowrap bgcolor="#EBF2F9" align="center"><input name="submitMove" type="submit" class="button" id="ת��" value="  ת��  "></td>
  </tr>
</form>
</table>
<br>
<%
End Function
'����ת�ƽڵ�------------------------------
Function saveMoveFolder()
  Dim rs,sql,fromID,fromParentID,fromSortPath,toID,toParentID,toSortPath,fromParentSortPath
  fromID=request.Form("ID")
  fromParentID=request.Form("ParentID")
  fromSortPath=request.Form("SortPath")
  toID=request.Form("toID")
  toParentID=request.Form("toParentID")
  toSortPath=request.Form("toSortPath")

  if toID="" or toParentID="" or toSortPath="" then
    response.write ("<script language=javascript> alert('û��ѡ���ƶ���Ŀ��λ�ã��뷵��ѡ��');history.back(-1);</script>")
    response.end
  end if
  if fromParentID=0 then
    response.write ("<script language=javascript> alert('һ�����಻�ܱ��ƶ����뷵��ѡ��');history.back(-1);</script>")
    response.end
  end if
  if fromSortPath=toSortPath then
    response.write ("<script language=javascript> alert('ѡ����ƶ�����Ŀ��λ����ͬ�ˣ��뷵������ѡ��');history.back(-1);</script>")
    response.end
  end if
  if Instr(toSortPath,fromSortPath)>0 or fromParentID=toID then
    response.write ("<script language=javascript> alert('���ܽ�����ƶ����������������뷵������ѡ��');history.back(-1);</script>")
    response.end
  end if
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "& Tbs &" where ID="&fromParentID
  rs.open sql,conn,0,1
  fromParentSortPath=rs("SortPath")
  conn.execute("update "& Tbs &" set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")'�����������·��
  conn.execute("update "& Tbs &" set ParentID='"&toID&"' where ID="&fromID)'���������ID
  conn.execute("update "& Tb &" set SortPath='"&toSortPath&"'+Mid(SortPath,Len('"&fromParentSortPath&"')+1) where Instr(SortPath,'"&fromSortPath&"')>0")'������Ϣ����·��
  response.write ("<script language=javascript> alert('�ƶ����ɹ������ȷ���鿴�������');location.replace('Sort.asp?TbS="&TbS&"&Tb="&Tb&"');</script>")
End Function
%>