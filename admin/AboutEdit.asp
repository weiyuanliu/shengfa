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
<TITLE>�༭��ҵ</TITLE>
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
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
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
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��ҵ��Ϣ����ӣ��޸Ľ�����ҵ��ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AboutEdit.asp?Result=Add" onClick='changeAdminFlag("�����ҵ��Ϣ")'>�����ҵ��Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="AboutList.asp" onClick='changeAdminFlag("��ҵ��Ϣ")'>�鿴��ҵ��Ϣ</a></td>
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
        <td height="20" align="right">��Ϣ���ƣ�</td>
        <td><input name="AboutName" type="text" class="textfield" id="AboutName" style="WIDTH: 240;" value="<%=AboutName%>" maxlength="100">&nbsp;������<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag then response.write ("checked")%>>
&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">�ꡡ���ǣ�</td>
        <td><input name="ChildFlag" type="checkbox" value="1" style='HEIGHT: 13px;WIDTH: 13px;' <%if ChildFlag then response.write ("checked")%>>&nbsp;��ҳ</td>
      </tr>
      <tr>
        <td height="20" align="right">�鿴Ȩ�ޣ�</td>
        <td><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> ����<input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��</td>
      </tr>
      <tr>
        <td height="20" rowspan="2" align="right" valign="top">��Ϣ���ݣ�<br>
        
        <td style="padding:6px"><textarea name="Content" rows="30" class="textfield" id="Content" style="WIDTH: 86%;"  ><%=Content%></textarea></td>
      </tr>
      <tr>
        <td></td>
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
sub AboutEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("AboutName")))="" then
      response.write ("<script language=javascript> alert('��Ϣ����Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '������վ����Ա
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
      GroupIdName=split(Request.Form("GroupID"),"���橾")
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
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_About where ID="&ID
      rs.open sql,conn,1,3
      rs("AboutName")=trim(Request.Form("AboutName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("Content")=Request.Form("Content")
      GroupIdName=split(Request.Form("GroupID"),"���橾")
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
    response.write "<script language=javascript> alert('�ɹ��༭��ҵ��Ϣ��');changeAdminFlag('��ҵ��Ϣ�б�');location.replace('AboutList.asp');</script>"
  else '��ȡ����Ա��Ϣ
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
