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
<TITLE>��ˡ��޸ġ��ظ�����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
call CreateEditor("Content")
%>


<%
if Instr(session("AdminPurview"),"|90,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result,ID
Result=request.QueryString("Result")
dim Msg_Name,Msg_Time,Msg_TelPhone,Replay,ReplayTime
ID=request.QueryString("ID")
call MesEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<STRONG>������Ϣ����ˣ��޸ģ��ظ�������Ϣ��ص�����</STRONG></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="MessageList.asp" onClick='changeAdminFlag("������Ϣ�б�")'>�鿴������Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp#Message" target="mainFrame" onClick='changeAdminFlag("��վ��Ϣ����")'>�����Ƿ��Զ����</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="RepalyMsg.asp?Action=SaveEdit&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">�����ߣ�</td>
        <td><input name="Msg_Name" type="text" class="textfield" id="Msg_Name" style="WIDTH: 240;" value="<%=Msg_Name%>">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">��������ϵ�绰��</td>
        <td><input name="Msg_TelPhone" type="text" class="textfield" id="Msg_TelPhone" style="WIDTH: 240;" value="<%=Msg_TelPhone%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">����ʱ�䣺</td>
        <td><input name="Msg_Time" type="text" class="textfield" id="Msg_Time" style="WIDTH: 240" value="<%=Msg_Time%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">�ظ�ʱ�䣺</td>
        <td><input name="ReplayTime" type="text" class="textfield" id="ReplayTime" style="WIDTH: 240" value="<%if ReplayTime<>"" then response.Write(ReplayTime) else response.Write(now())%>" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">�ظ����ݣ�</td>
        <td>
        <textarea name="Content" rows="15" class="textfield" id="Content" style="WIDTH: 100%;"  >
        	<%if Replay="" or isnull(Replay) then%>
            	<%=Msg_Name%>���ѣ����Ļ�����<font color="#ff0000"><u> x��x��x�� </u></font>���������Ŀ�ݵ�����<font color="#ff0000"><u> xxxxxxxxxx </u></font>,��������Ϊ�����Ͳ�Ʒ�Ŀ�ݹ�˾����ϵ�绰<font color="#ff0000"><u> xxxxxxxx </u></font>����ݹ�˾��<font color="#ff0000"><u> xxx��ݹ�˾ </u></font>���뼰ʱ���ݹ�˾��ϵ�������߿�ݹ�˾���Ŀ�ݵ��ţ������Ǽ�ʱ�ͻ���������������������µ�400-661-9668�����ǹ�����Ա���������
            <%else%>
            	<%=Replay%>
            <%end if%>
        </textarea>
       </td>
      </tr>
	  <tr>
	  <td height="40" align="right" >����Ա��</td>
	  <td><%=session("UserName")%></td>
	  </tr>
      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="����" style="WIDTH: 80;" >
          <input name="submitSaveEdit2" type="button" class="button"  id="submitSaveEdit2" value="����" style="WIDTH: 80;" onClick="window.location.href=document.referrer;" ></td>
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
sub MesEdit()
  if ID=""or isnull(ID) or not(IsNumeric(ID)) then
  	response.Write("<script langauge=javascript>"&vbcrlf)
		response.Write("alert('�Բ������ݳ����뷵�أ�');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>"&vbcrlf)
	response.End()
  end if
  Dim Action,Rs,Sql,Editadmin
  Editadmin=session("UserName")
  Action=Trim(Request.QueryString("Action"))
  Set rs=server.CreateObject("adodb.recordset")
  sql="select * from MsgData where id="&id
  if Action="SaveEdit" then
  	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('��¼δ�ҵ����뷵�أ�');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	else
		rs("Replay")=Trim(Request.Form("Content"))
		rs("ReplayTime")=Trim(Request.Form("ReplayTime"))
		rs("Ediadmin")=Editadmin
		rs.update()
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('�ظ����Գɹ���');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
	end if
  else
  	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('��¼δ�ҵ����뷵�أ�');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
  	else
		Msg_Name=rs("Msg_Name")
		Msg_Time=rs("Msg_Time")
		Msg_TelPhone=rs("Msg_TelPhone")
		Replay=rs("Replay")
		ReplayTime=rs("ReplayTime")
		Editadmin=rs("Ediadmin")
		rs.close()
		set rs=Nothing
  	end if
  end if
end sub

%>