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
<TITLE>�鿴���޸ġ��ظ�����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|305,")=0 then 
  response.write ("<script language=javascript> alert('�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�');history.back(-1);</script>")
end if
%>
<%
if Instr(session("AdminPurview"),"|305,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ReplyContent,ReplyTime,ID,ProductName,ProductNo,Amount,Remark,display,NotSend
dim Linkman,Company,Address,ZipCode,Telephone,Fax,Mobile,Email,AddTime,States,FuKuan,HuoDao_FuKuan,Tel
ID=request.QueryString("ID")
Dim OrderSate:OrderSate=Cll()
if OrderSate="" then OrderSate="δ����|�����Ѹ�|Ǯ���ѷ�|���ܵ���|�Ѿ�����"
Function Cll()
 Dim rs,sql
 set rs=server.CreateObject("Adodb.recordset")
 sql="Select top 1 OrderSates from NwebCn_Site"
 rs.open sql,conn,1,1
 if not rs.eof then
 Cll=rs(0)
 end if
End Function
call OrderEdit() 

%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>������Ϣ���鿴���޸ģ��ظ�������Ϣ��ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="OrderList.asp" onClick='changeAdminFlag("������Ϣ�б�")'>�鿴������Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp" target="mainFrame" onClick='changeAdminFlag("��վ��Ϣ����")'>��վ��Ϣ����</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="OrderEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="160" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">��Ʒ���ƣ�</td>
        <td><input name="ProductName" type="text" class="textfield" id="ProductName" style="WIDTH: 240;" value="<%=ProductName%>">&nbsp;&nbsp;
        	<%if HuoDao_FuKuan then%>
            <input type="hidden" name="HuoDao_FuKuan" id="HuoDao_FuKuan" value="1">
        	<input name="FuKuan" id="FuKuan" type="checkbox" value="1" <%if FuKuan then response.Write("Checked")%>>&nbsp;�����󸶿�
            <%else%>
            <input type="hidden" name="HuoDao_FuKuan" id="HuoDao_FuKuan" value="0">
        	<%end if%>
        </td>
      </tr>
      <tr>
        <td height="20" align="right">��Ʒ��ţ�</td>
        <td><input name="ProductNo" type="text" class="textfield" id="ProductNo" style="WIDTH: 240;" value="<%=ProductNo%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">����������</td>
        <td><input name="Amount" type="text" class="textfield" id="Amount"   value="<%=Amount%>" size="80">��ʽ����"�����һ��(XX)|����Ӷ���(XX)"��д ������� </td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">����˵����
        <td><textarea name="Remark" rows="6" class="textfield" id="Remark" style="WIDTH: 76%;"><%=PringText(Remark)%></textarea></td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;��&nbsp;�ߣ�</td>
        <td><%=Linkman%></td>
      </tr>
      <tr>
        <td height="20" align="right">��λ���ƣ�</td>
        <td><input name="Company" type="text" class="textfield" id="Company" style="WIDTH: 240;" value="<%=Company%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">ͨ�ŵ�ַ��</td>
        <td><input name="Address" type="text" class="textfield" id="Address"  value="<%=Address%>" size="80"></td>
      </tr>
      <tr>
        <td height="20" align="right">�������ţ�</td>
        <td><input name="ZipCode" type="text" class="textfield" id="ZipCode" style="WIDTH: 120" value="<%=ZipCode%>"></td>
      </tr>
<%
'���ε绰����
dim TelStr
if session("AdminId") = 62 or session("AdminId") = 1 then
	TelStr = Tel
else
	TelStr = Left(Tel,0)&"********"&right(Tel,3)
end if
%>
      <tr>
        <td height="20" align="right">�硡������</td>
        <td><input name="Telephonepb" type="text" class="textfield" id="Telephonepb" style="WIDTH: 240;" value="<%=TelStr%>">
		<input name="Telephone" type="hidden" id="Telephone" value="<%=Tel%>">
	</td>
      </tr>
      <tr>
        <td height="20" align="right">�������棺</td>
        <td><input name="Fax" type="text" class="textfield" id="Fax" style="WIDTH: 120" value="<%=Fax%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">�ƶ��绰��</td>
        <td><input name="Mobile" type="text" class="textfield" id="Mobile" style="WIDTH: 120" value="<%=Mobile%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">�������䣺</td>
        <td><input name="Email" type="text" class="textfield" id="Email" style="WIDTH: 240" value="<%=Email%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">����ʱ�䣺</td>
        <td><input name="AddTime" type="text" class="textfield" id="AddTime" style="WIDTH: 240" value="<%=AddTime%>"></td>
      </tr>
      <tr>
        <td height="20" align="right">�޸Ķ���״̬��</td>
        <td valign="bottom"><%
    response.Write("<select name='State' id='State' size='1' style='margin-left:10px;' onchange='Event_Chang();'>")
  	response.Write("<option value='NULL'>--������--</option>")
	
	Dim S:S=split(OrderSate,"|")
	Dim i
	for i=lbound(S) to ubound(S)
	 %>
	 <option value="<%=S(i)%>" <%if S(i)=states then response.Write("selected")%>><%=S(i)%></option>
	 <%
	next
	if states="���ܵ���" then
	display=""
else
display="none"
	end if
	
	'if states="�����󸶿�" then
		'response.Write("<option value='�����󸶿�' selected>�����󸶿�</option>")
	'else
		'response.Write("<option value='�����󸶿�'>�����󸶿�</option>")
	'end if
	'if states="���ܷ���" then
		'response.Write("<option value='���ܷ���' selected>���ܷ���</option>")
	'else
		'response.Write("<option value='���ܷ���'>���ܷ���</option>")
	'end if
  response.Write("</select>")%>
         <script language="javascript">
		 	<!--
			
			
			function Event_Chang()
			{
				var Stats,NotSend;
				Stats=document.getElementById("State");
				NotSend=document.getElementById("NotSend");
				if((Stats.value).indexOf("�ܵ���")!=-1 || (Stats.value).indexOf("���ܷ���")!=-1)
				{
					NotSend.style.display="";
				}
				else
				{
					NotSend.style.display="none";
				}
			}
			-->
		 </script>
         <span style="margin-left:20px;display:<%=display%>;" id="NotSend">
         	<input type="text" name="NotSend" size="50" value="<%=NotSend%>"/>&nbsp;<font color="#FF0000">*����дԭ��</font>
         </span>
       </td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom"><label>
          <input type="submit" name="Modify" id="Modify" value="�� ��">
          <input type="button" name="Modify2" id="Modify2" value="�� ��" onClick="window.history.go(-1);">
        </label></td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
</BODY>
</HTML>
<%
sub OrderEdit()
  dim Action,rsCheckAdd,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Order where ID="&ID
      rs.open sql,conn,1,3
	  
	 ' if trim(Request.Form("HuoDao_FuKuan"))="1" then
		  'if Trim(Request.Form("FuKuan"))="1" then
			 ' rs("FuKuan")=true
			 ' rs("State")=StrReplace(Request.Form("Stats"))
		  'else
			 ' rs("FuKuan")=false
		 ' end if
	 ' else
	  	'rs("State")=StrReplace(Request.Form("Stats"))	  
	 ' end if
	  
	  rs("State")=StrReplace(Request.Form("State"))	  
	  Rs("Amount")=Trim(Request.Form("Amount"))
	 ' response.Write(Replace(Replace(Replace(Replace(Trim(Request.Form("Remark")),"֧����ʽ��","|"),"Ӧ����","|"),"�ͻ���ʽ��",""),vbcrlf,""))
	 ' Response.End()
	  Rs("Remark")=Replace(Replace(Replace(Replace(Trim(Request.Form("Remark")),"֧����ʽ��","|"),"Ӧ����","|"),"�ͻ���ʽ��",""),vbcrlf,"")
	  rs("Company")=Trim(Request.Form("Company"))
	  rs("Address")=Trim(Request.Form("Address"))
	  rs("ZipCode")=Trim(Request.Form("ZipCode"))
	  rs("Tel")=Trim(Request.Form("Telephone"))
	  rs("Fax")=Trim(Request.Form("Fax"))
	  rs("Telephone")=Trim(Request.Form("Mobile"))
	  rs("Email")=Trim(Request.Form("Email"))
	  rs("AddTime")=Trim(Request.Form("AddTime"))
	  if Trim(Request.Form("NotSend"))<>"" then
	  	rs("NotSendText")=trim(Request.Form("NotSend"))
	  end if
	  if instr(Trim(Request.Form("Stats")),"���ѷ�")>0 then
	  	rs("FaHuoTime")=Now()
	  end if
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭������Ϣ��');changeAdminFlag('������Ϣ�б�');location.replace('OrderList.asp');</script>"
  else '��ȡ������Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Order where ID="& ID
      rs.open sql,conn,1,1
	  
	  ProductName=rs("ProductName")
	  ProductNo=rs("ProductNo")
	  Amount=rs("Amount")
	  Remark=ReStrReplace(rs("Remark"))
	  Linkman=GuestInfo(rs("MemID"),rs("Linkman"),rs("Sex"))
	  Company=rs("Company")
	  Address=rs("Address")
	  ZipCode=rs("ZipCode")
	  FuKuan=rs("FuKuan")
	  States=rs("State")
	  NotSend=rs("NotSendText")
	  Tel=rs("Tel")
	  Fax=rs("Fax")
	  Mobile=rs("Telephone")
	  Email=rs("Email")
	  HuoDao_FuKuan=rs("HuoDao_FuKuan")
	  AddTime=rs("AddTime")
	  ReplyContent=ReStrReplace(rs("ReplyContent"))
	  ReplyTime=rs("ReplyTime")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub

function GuestInfo(ID,Guest,Sex)
  'Dim rs,sql
  'Set rs=server.CreateObject("adodb.recordset")
  'sql="Select * From NwebCn_Members where ID="&ID
  'rs.open sql,conn,1,1
  'if rs.bof and rs.eof then
    GuestInfo=Guest & "&nbsp;" & Sex
  'else
    'GuestInfo="<font color='green'>��Ա&nbsp;</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""ǰ̨��Ա����"")'>"&Guest&"</a>"&Sex
  'end if
  'rs.close
  'set rs=nothing
end function 

function Print(Amount)
	dim str,i,str1,str2,str3
	str1=""
	str=split(Amount,"|")
	for i=0 to ubound(str)
		if i>0 then str1=str1&"��"
		if str1="" then
			str1=Mid(str(i),1,instr(str(i),"(")-1)
			str2=Mid(str(0),instr(str(i),"(")+1,1)
			str3=""
		else
			str1=str1&Mid(str(i),1,instr(str(i),"(")-1)
			str2=Mid(str(0),instr(str(i),"(")+1,1)
			str3=Mid(str(1),instr(str(i),"(")+1,1)
		end if
		str1=str1&Mid(str(i),instr(str(i),"(")+1,(instr(str(i),")"))-(instr(str(i),"(")+1))&"��"
	next
	Print=str1&"||"&str2&"||"&str3
end function

function PringText(Remark)
	dim str,str1,i
	str=split(Remark,"|")
	if ubound(str)>0 then
	str1="�ͻ���ʽ��"&str(0)
	str1=str1&vbcrlf
	str1=str1&"֧����ʽ��"&str(1)
	str1=str1&vbcrlf
	str1=str1&"Ӧ����"&str(2)
	PringText=str1
	end if
end function
%>