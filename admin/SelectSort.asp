<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'����������������������������������������������������������������
'����������������������������������������������������������������
'�������������������տƼ���ҵ��վ����ϵͳ��LISuo����������������  ��
'����������������������������������������������������������������
' ����Ȩ���С�qisehu.com
'
'�����������������տƼ����޹�˾
'��������������Add:�Ĵ�ʡ�ɶ��ж���·������181��13¥20/21��
'����������������������������������������������������������������
'����������������������������������������������������������������
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>ѡ�����</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->

<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
    <tr>
      <td height="24" nowrap><A href="#Help"><IMG height="18" alt="ʹ�ð�����Help��" src="Images/Explain.gif" width="18" align="absMiddle" border="0"></A><font color="#FFFFFF"><strong>ѡ��������𣺵��'ѡ��'���</strong></font></td>
    </tr>
    <tr>
      <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" height="6" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td></td>
        </tr>
      </table>
	  <%
      Dim Result,Datafrom
      Result=request.QueryString("Result")
      Select Case Result
        Case "Products"
		  Datafrom="NwebCn_ProductSort"
        Case "News"
		  Datafrom="NwebCn_NewsSort"
        Case "Download"
		  Datafrom="NwebCn_DownSort"
        Case "Need"
		  Datafrom="NwebCn_NeedSort"
        Case "Others"
		  Datafrom="NwebCn_OthersSort"
        Case Else
		
      End Select
	  ListSort(0)
      %>
	  </td>
    </tr>
</table>
</BODY>
</HTML>
<%
Function ListSort(id)
  Dim rs,sql,i,ChildCount,FolderType,FolderName,onMouseUp,ListType
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "&Datafrom&" where ParentID="&id&" order by id"
  rs.open sql,conn,1,1
  if id=0 and rs.recordcount=0 then
    response.write ("���޷���!")
    response.end
  end if  
  i=1
  response.write("<table border='0' cellspacing='0' cellpadding='0'>")
  while not rs.eof
    ChildCount=conn.execute("select count(*) from "&Datafrom&" where ParentID="&rs("id"))(0)
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
 response.write("<a href=javaScript:AddSort('"&escape(SortText(rs("ID")))&"','"&rs("ID")&"','"&rs("SortPath")&"')><font color='#ff6600'>ѡ��</font></a>")
    response.write("</td></tr>")
    if ChildCount>0 then
%>
      <tr id="a<%= rs("id")%>" style="display:yes"><td class="<%= ListType%>" nowrap></td><td ><% ListSort(rs("id")) %></td></tr>
<%
    end if
    rs.movenext
    i=i+1
  wend
  response.write("</table>")
  rs.close
  set rs=nothing
end function
%>
<%
'�����������--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From "&Datafrom&" where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>