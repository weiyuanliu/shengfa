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
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>�༭��Ƹ</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|98,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,JobName,ViewFlag,JobAddress,JobNumber,Emolument,EndDate,Content,px
ID=request.QueryString("ID")
call JobsEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��Ƹ��Ϣ����ӣ��޸���Ƹ����ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="JobsEdit.asp?Result=Add" onClick='changeAdminFlag("�����Ƹ��Ϣ")'>�����Ƹ��Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="JobsList.asp" onClick='changeAdminFlag("��Ƹ��Ϣ�б�")'>�鿴��Ƹ��Ϣ</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="JobsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">ְλ���ƣ�</td>
        <td><input name="JobName" type="text" class="textfield" id="JobName" style="WIDTH: 240;" value="<%=JobName%>">&nbsp;������<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag or Result="Add" then response.write ("checked")%>>&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">�����ص㣺</td>
        <td><input name="JobAddress" type="text" class="textfield" id="JobAddress" style="WIDTH: 240;" value="<%=JobAddress%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">��Ƹ������</td>
        <td><input name="JobNumber" type="text" class="textfield" id="JobNumber" style="WIDTH: 240" value="<%=JobNumber%>">&nbsp;*&nbsp;6��</td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;н&nbsp;ˮ��</td>
        <td><input name="Emolument" type="text" class="textfield" id="Emolument" style="WIDTH: 240;" value="<%=Emolument%>">&nbsp;*&nbsp;3000Ԫ/��</td>
      </tr>
	        <tr>
        <td height="20" align="right">��&nbsp;��</td>
        <td><input name="Px" type="text" class="textfield" id="Px" style="WIDTH: 240;" value="<%=px%>">&nbsp;*&nbsp;ֻ����д����</td>
      </tr>
      <tr>
        <td height="20" align="right">�������ڣ�</td>
        <td><input name="EndDate" type="text" class="textfield" id="EndDate" style="WIDTH: 240;" value="<% if EndDate="" then response.write (DateAdd("m",3,now())) else response.write (EndDate) end if%>" maxlength="14">&nbsp;*&nbsp;Ĭ��Ϊ3���£����ֶ��������ڸ�ʽ</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">��Ϣ���ݣ�<br>
		  <img title="���������ӻ��鿴���༭����..." src="Images/Edit.gif" width="51" height="20" style="cursor:hand" onClick="OpenDialog('../Editor/EditorDialog.html?lnk=Content&file=Editor_1.html',800,520);">
        <td><!--<textarea name="Content" rows="12" class="textfield" id="Content" style="WIDTH: 86%;" readonly><%=Content%></textarea>-->

            <textarea name="Content" rows="12" class="textfield" id="Content" style="WIDTH: 86%;" ><%=Content%></textarea><br>

��������"&lt;br&gt;"����,�����ñ༭���༭</td>
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
sub JobsEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("JobName")))<3 then
      response.write ("<script language=javascript> alert('ְλ����Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
    if len(trim(request.Form("JobAddress")))="" or len(trim(request.Form("JobNumber")))="" or len(trim(request.Form("Emolument")))="" then
      response.write ("<script language=javascript> alert('""�����ص㡢ְλ��������нˮ""����Ϊ������Ŀ���Ҳ�����2���ַ���');history.back(-1);</script>")
      response.end
    end if
    if len(trim(request.Form("EndDate")))<4 then
      response.write ("<script language=javascript> alert('""��������""����Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '������վ����Ա
	  sql="select * from NwebCn_Jobs"
      rs.open sql,conn,1,3
      rs.addnew
      rs("JobName")=trim(Request.Form("JobName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("JobAddress")=trim(Request.Form("JobAddress"))
	  rs("JobNumber")=trim(Request.Form("JobNumber"))
	  rs("Emolument")=trim(Request.Form("Emolument"))
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  rs("EndDate")=trim(Request.Form("EndDate"))
	  rs("Content")=Request.Form("Content")
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Jobs where ID="&ID
      rs.open sql,conn,1,3
      rs("JobName")=trim(Request.Form("JobName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  rs("JobAddress")=trim(Request.Form("JobAddress"))
	  rs("JobNumber")=trim(Request.Form("JobNumber"))
	  rs("Emolument")=trim(Request.Form("Emolument"))
	 
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  rs("EndDate")=trim(Request.Form("EndDate"))
	  rs("Content")=Request.Form("Content")
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭��Ƹ��Ϣ��');changeAdminFlag('��Ƹ��Ϣ�б�');location.replace('JobsList.asp');</script>"
  else '��ȡ����Ա��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Jobs where ID="& ID
      rs.open sql,conn,1,1
	  JobName=rs("JobName")
	  ViewFlag=rs("ViewFlag")
	  JobAddress=rs("JobAddress")
	  JobNumber=rs("JobNumber")
	  Emolument=rs("Emolument")
	  px=Rs("Px")
	  EndDate=rs("EndDate")	  
      Content=rs("Content")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
%>