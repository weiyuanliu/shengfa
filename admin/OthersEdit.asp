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
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>�༭������Ϣ</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|73,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,OthersName,ViewFlag,SortName,SortID,SortPath,px
dim GroupID,GroupIdName,Exclusive,Content
ID=request.QueryString("ID")
call OthersEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>���ż���������鿴����ӣ��޸ģ�ɾ��������Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="OthersEdit.asp?Result=Add" onClick='changeAdminFlag("��Ӳ�Ʒ��Ϣ")'>���������Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="OthersList.asp" onClick='changeAdminFlag("��Ʒ�б�")'>�鿴����������Ϣ</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="OthersEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editOthers idth="100%">
      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">��Ϣ���ƣ�</td>
        <td><input name="OthersName" type="text" class="textfield" id="OthersName" style="WIDTH: 240;" value="<%=OthersName%>" maxlength="100">&nbsp;������<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag then response.write ("checked")%>>&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">�������</td>
        <td><input name="SortName" type="text" class="textfield" id="SortName" value="<%=SortName%>" style="WIDTH: 240;background-color:#EBF2F9;" readonly>&nbsp;<a href="javaScript:OpenScript('SelectSort.asp?Result=Others',500,500,'')"><img src="Images/Select.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">������֣�</td>
        <td><input name="SortID" type="text" class="textfield" id="SortID" style="WIDTH: 40;background-color:#EBF2F9;" value="<%=SortID%>" readonly><input name="SortPath" type="text" class="textfield" id="SortPath" style="WIDTH: 200;background-color:#EBF2F9;" value="<%=SortPath%>" readonly>&nbsp;*</td>
      </tr>
	        <tr>
        <td height="20" align="right">����</td>
        <td><input name="px" type="text" class="textfield" id="px" value="<%=px%>" >&nbsp;*ֻ����д����</td>
      </tr>
      <tr>
        <td height="20" align="right">�鿴Ȩ�ޣ�</td>
        <td><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> ����<input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">��ϸ���ݣ�<br>
		  <img title="���������ӻ��鿴���༭����..." src="Images/Edit.gif" width="51" height="20" style="cursor:hand" onClick="OpenDialog('../Editor/EditorDialog.html?lnk=Content&file=Editor_1.html',800,520);">
        <td><textarea name="Content" rows="12" class="textfield" id="Content" style="WIDTH: 86%;" readonly><%=Content%></textarea></td>
      </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="����" style="WIDTH: 80;" ></td>
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
sub OthersEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭��Ʒ��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("OthersName")))<3 then
      response.write ("<script language=javascript> alert('��Ϣ����Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '������Ʒ��Ϣ
	  sql="select * from NwebCn_Others"
      rs.open sql,conn,1,3
      rs.addnew
      rs("OthersName")=trim(Request.Form("OthersName"))
	  if isnumeric(request.Form("px")) then
	  rs("Px")=Request.Form("Px")
	  else
	   rs("Px")=0
	  end if 
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")="" and Request.Form("SortPath")="" then
        response.write ("<script language=javascript> alert('��ѡ���������࣡');history.back(-1);</script>")
        response.end
	  else
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("Content")=trim(Request.Form("Content"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸Ĳ�Ʒ��Ϣ
      sql="select * from NwebCn_Others where ID="&ID
      rs.open sql,conn,1,3
      rs("OthersName")=trim(Request.Form("OthersName"))
	   if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('��ѡ���������࣡');history.back(-1);</script>")
        response.end
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("Content")=trim(Request.Form("Content"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭������Ϣ��');changeAdminFlag('������Ϣ�б�');location.replace('OthersList.asp');</script>"
  else '��ȡ��Ʒ��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Others where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("���ݿ��ȡ��¼����")
        response.end
      end if
	  OthersName=rs("OthersName")
	  px=rs("px")
	  ViewFlag=rs("ViewFlag")
	  SortName=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  SortPath=rs("SortPath")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
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
    response.write("δ�����")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"���橾"&rs("GroupName")&"'")
    if GroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
end sub
%>
<%
'�����������--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_OthersSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>