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
<TITLE>�鿴���޸ġ��ظ���Ӧ��Ϣ</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|96,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ReplyContent,ReplyTime,ID,NeedID,SupplyName,Remark
dim Linkman,Company,Address,ZipCode,Telephone,Fax,Mobile,Email,AddTime
ID=request.QueryString("ID")
call SupplyEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��Ӧ��Ϣ���鿴���޸ģ��ظ���Ӧ��Ϣ��ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="SupplyList.asp" onClick='changeAdminFlag("��Ӧ��Ϣ�б�")'>�鿴��Ӧ��Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("��վ��Ϣ����")'>��վ��Ϣ����</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="SupplyEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">��Ӧ��Ʒ��</td>
        <td><input name="SupplyName" type="text" class="textfield" id="SupplyName" style="WIDTH: 240;" value="<%=SupplyName%>" readonly>&nbsp;<a href=NeedEdit.asp?Result=Modify&ID=<%=NeedID%> target="mainFrame" onClick='changeAdminFlag("�鿴������Ϣ")'>�鿴����</a></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">����˵����
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 76%;" readonly><%=Remark%></textarea></td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;Ӧ&nbsp;�ߣ�</td>
        <td><%=Linkman%></td>
      </tr>
      <tr>
        <td height="20" align="right">��λ���ƣ�</td>
        <td><input name="Company" type="text" class="textfield" id="Company" style="WIDTH: 240;" value="<%=Company%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">ͨ�ŵ�ַ��</td>
        <td><input name="Address" type="text" class="textfield" id="Address" style="WIDTH: 240;" value="<%=Address%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">�ʡ����ࣺ</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="WIDTH: 120" value="<%=ZipCode%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">�硡������</td>
        <td><input name="Telephone" type="text" class="textfield" id="Telephone" style="WIDTH: 240;" value="<%=Telephone%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">�������棺</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="WIDTH: 120" value="<%=Fax%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">�ƶ��绰��</td>
        <td><input name="Mobile" type="text" class="textfield" id="Mobile" style="WIDTH: 120" value="<%=Mobile%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">�������䣺</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="WIDTH: 240" value="<%=Email%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">����ʱ�䣺</td>
        <td><input name="AddTime" type="text" class="textfield" id="AddTime" style="WIDTH: 240" value="<%=AddTime%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">�ظ�ʱ�䣺</td>
        <td><input name="ReplyTime" type="text" class="textfield" id="ReplyTime" style="WIDTH: 240" value="<%=ReplyTime%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">�ظ����ݣ�</td>
        <td><textarea name="ReplyContent" rows="6" class="textfield" id="ReplyContent" style="WIDTH: 76%;"><%=ReplyContent%></textarea></td>
      </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="����" style="WIDTH: 80;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom"></td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub SupplyEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����ظ���Ӧ��Ϣ
    set rs = server.createobject("adodb.recordset")
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Supply where ID="&ID
      rs.open sql,conn,1,3
	  rs("ReplyContent")=StrReplace(Request.Form("ReplyContent"))
	  if not trim(request.Form("ReplyContent"))="" then
	    rs("ReplyTime")=now()
      end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭���ظ���Ӧ��Ϣ��');changeAdminFlag('��Ӧ��Ϣ�б�');location.replace('SupplyList.asp');</script>"
  else '��ȡ������Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Supply where ID="& ID
      rs.open sql,conn,1,1
	  NeedID=rs("NeedID")
	  SupplyName=rs("SupplyName")
	  Remark=ReStrReplace(rs("Remark"))
	  Linkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  Company=rs("Company")
	  Address=rs("Address")
	  ZipCode=rs("ZipCode")
	  Telephone=rs("Telephone")
	  Fax=rs("Fax")
	  Mobile=rs("Mobile")
	  Email=rs("Email")
	  AddTime=rs("AddTime")
	  ReplyContent=ReStrReplace(rs("ReplyContent"))
	  ReplyTime=rs("ReplyTime")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub

function GuestInfo(ID,Guest,Sex)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    GuestInfo=Guest & "&nbsp;" & Sex
  else
    GuestInfo="<font color='green'>��Ա&nbsp;</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""ǰ̨��Ա����"")'>"&Guest&"</a>"&Sex
  end if
  rs.close
  set rs=nothing
end function 
%>