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
<TITLE>�༭��Ա</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|103,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,MemName,RealName,Password,vPassword,Sex,GroupID,GroupName,GroupIdName
dim Company,Address,ZipCode,Telephone,Fax,Mobile,Email,Homepage,Working
ID=request.QueryString("ID")
call MemEdit() 
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��վ��Ա����ӣ��޸Ļ�Ա��Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="MemEdit.asp?Result=Add" onClick='changeAdminFlag("����»�Ա")'>����»�Ա</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="MemList.asp" onClick='changeAdminFlag("�鿴���л�Ա")'>�鿴���л�Ա</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editMemForm" method="post" action="MemEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onSubmit="return CheckMemEdit()">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;¼&nbsp;����</td>
        <td><input name="MemName" type="text" class="textfield" id="MemName" style="WIDTH: 120;" value="<%=MemName%>" maxlength="16" <%if Result="Modify" then response.write ("readonly")%>>&nbsp;*&nbsp;3-16λ�ַ��������޸�</td>
      </tr>
      <tr>
        <td height="20" align="right">��ʵ������</td>
        <td><input name="RealName" type="text" class="textfield" id="RealName" style="WIDTH: 120;" value="<%=RealName%>" maxlength="16"></td>
      </tr>
      <tr>
        <td height="20" align="right">�ܡ����룺</td>
        <td><input name="Password" type="password" class="textfield" id="Password" maxlength="20" style="WIDTH: 120;">&nbsp;*&nbsp;6-16λ�ַ��������δ�޸�����</td>
      </tr>
      <tr>
        <td height="20" align="right">ȷ�����룺</td>
        <td><input name="vPassword" type="password" class="textfield" id="vPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">�ԡ�����</td>
        <td><input type="radio" name="sex" value="����" <%if Sex="����" then response.write ("checked")%>>&nbsp;����&nbsp;<input type="radio" name="sex" value="Ůʿ" <%if Sex="Ůʿ" then response.write ("checked")%>>&nbsp;Ůʿ</td>
      </tr>
      <tr>
        <td height="20" align="right">��Ա���</td>
        <td>
		<select name="GroupID" class="textfield"><% call SelectGroup() %>
        </select></td>
      </tr>
      <tr>
        <td height="20" align="right">��λ���ƣ�</td>
        <td><input name="Company" type="text" class="textfield" id="Company" style="WIDTH: 240;" value="<%=Company%>" maxlength="100"></td>
      </tr>
      <tr>
        <td height="20" align="right">�ء���ַ��</td>
        <td><input name="Address" type="text" class="textfield" id="Address" style="WIDTH: 240;" value="<%=Address%>" maxlength="100"></td>
      </tr>
      <tr>
        <td height="20" align="right">�ʡ����ࣺ</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="WIDTH: 120;" value="<%=ZipCode%>" maxlength="16"></td>
      </tr>
      <tr>
        <td height="20" align="right">�硡������</td>
        <td><input name="Telephone" type="text" class="textfield" id="Telephone" style="WIDTH: 240;" value="<%=Telephone%>" maxlength="50"></td>
      </tr>
      <tr>
        <td height="20" align="right">�������棺</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="WIDTH: 120;" value="<%=Fax%>" maxlength="16"></td>
      </tr>
      <tr>
        <td height="20" align="right">�ƶ��绰��</td>
        <td><input name="Mobile" type="text" class="textfield" id="Mobile" style="WIDTH: 120;" value="<%=Mobile%>" maxlength="16"></td>
      </tr>
      <tr>
        <td height="20" align="right">�������䣺</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="WIDTH: 240;" value="<%=Email%>" maxlength="50"></td>
      </tr>
      <tr>
        <td height="20" align="right">������ַ��</td>
        <td><input name="HomePage" type="text" class="textfield" id="HomePage" style="WIDTH: 240;" value="<%=HomePage%>" maxlength="50"></td>
      </tr>
      <tr>
        <td height="20" align="right">������Ч��</td>
        <td><input name="Working" type="checkbox"  value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if Working then response.write ("checked")%>></td>
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
sub MemEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
    if Result="Add" then '������վ����Ա
      set rsRepeat = conn.execute("select MemName from NwebCn_Members where MemName='" & trim(Request.Form("MemName")) & "'")
      if not (rsRepeat.bof and rsRepeat.eof) then '�жϴ˹���Ա���Ƿ����
        response.write "<script language=javascript> alert('" & trim(Request.Form("MemName")) & "�˻�Ա���Ѿ����ڣ��뻻һ����¼�������ԣ�');history.back(-1);</script>"
        response.end
      end if 
	  sql="select * from NwebCn_Members"
      rs.open sql,conn,1,3
      rs.addnew
      rs("MemName")=trim(Request.Form("MemName"))
      rs("RealName")=StrReplace(trim(Request.Form("RealName")))
      if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
        response.write "<script language=javascript> alert('��Ա���������ַ���Ϊ6-16λ��');history.back(-1);</script>"
        response.end
      end if
	  if Request.Form("Password")<>Request.Form("vPassword") then 
        response.write "<script language=javascript> alert('������������벻һ����');history.back(-1);</script>"
        response.end
	  end if
	  rs("Password")=Md5(Request.Form("Password"))
	  rs("Sex")=Request.Form("Sex")
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("GroupName")=GroupIdName(1)
	  rs("Company")=StrReplace(trim(Request.Form("Company")))
	  rs("Address")=StrReplace(trim(Request.Form("Address")))
	  rs("ZipCode")=StrReplace(trim(Request.Form("ZipCode")))
	  rs("Telephone")=StrReplace(trim(Request.Form("Telephone")))
	  rs("Fax")=StrReplace(trim(Request.Form("Fax")))
	  rs("Mobile")=StrReplace(trim(Request.Form("Mobile")))
	  rs("Email")=trim(Request.Form("Email"))
	  rs("HomePage")=StrReplace(trim(Request.Form("HomePage")))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Members where ID="&ID
      rs.open sql,conn,1,3
      rs("MemName")=trim(Request.Form("MemName"))
      rs("RealName")=StrReplace(trim(Request.Form("RealName")))
      if trim(Request.Form("Password"))<>"" then
	    if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
          response.write "<script language=javascript> alert('��Ա���������ַ���Ϊ6-16λ��');history.back(-1);</script>"
          response.end
        end if
	    if Request.Form("Password")<>Request.Form("vPassword") then 
          response.write "<script language=javascript> alert('������������벻һ����');history.back(-1);</script>"
          response.end
	    end if
	    rs("Password")=Md5(Request.Form("Password"))
	  end if
	  rs("Sex")=Request.Form("Sex")
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("GroupName")=GroupIdName(1)
	  rs("Company")=StrReplace(trim(Request.Form("Company")))
	  rs("Address")=StrReplace(trim(Request.Form("Address")))
	  rs("ZipCode")=StrReplace(trim(Request.Form("ZipCode")))
	  rs("Telephone")=StrReplace(trim(Request.Form("Telephone")))
	  rs("Fax")=StrReplace(trim(Request.Form("Fax")))
	  rs("Mobile")=StrReplace(trim(Request.Form("Mobile")))
	  rs("Email")=StrReplace(trim(Request.Form("Email")))
	  rs("HomePage")=StrReplace(trim(Request.Form("HomePage")))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭��վ��Ա��');changeAdminFlag('���л�Ա');location.replace('MemList.asp');</script>"
  else '��ȡ����Ա��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Members where ID="& ID
      rs.open sql,conn,1,1
	  MemName=rs("MemName")
	  RealName=rs("RealName")
	  Sex=rs("Sex")
	  GroupID=rs("GroupID")
	  GroupName=rs("GroupName")
	  Company=rs("Company")
	  Address=rs("Address")
	  ZipCode=rs("ZipCode")
	  Telephone=rs("Telephone")
	  Fax=rs("Fax")
	  Mobile=rs("Mobile")
	  Email=rs("Email")
	  Homepage=rs("Homepage")
	  Working=rs("Working")
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
