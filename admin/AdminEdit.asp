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
<TITLE>�༭����Ա</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|101,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result,Action
Result=request.QueryString("Result")
dim ID,AdminName,Working,Password,vPassword,UserName,Purview,Explain,AddTime,GroupID,GroupName,GroupIdName,groupid1
ID=request.QueryString("ID")
groupid1=request.QueryString("GroupID")
if ID="" then ID=0
call AdminEdit() 
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��վ����Ա����ӣ��޸Ĺ���Ա��Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AdminEdit.asp?Result=Add" onClick='changeAdminFlag("��ӹ���Ա")'>��ӹ���Ա</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="AdminList.asp" onClick='changeAdminFlag("��վ����Ա")'>�鿴���й���Ա</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="AdminEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onSubmit="return CheckAdminEdit()">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	<table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">
	  <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  
	  <tr <% if ID=1 then Response.Write("style='display:none'")%> >
	    <td height="20" align="right">����Ա�飺</td>
		<%
		if Result ="Add" then
		%>
		<td><select name="GroupID" class="textfield"><% call SelectGroup() %></select></td>
		<%
		end if
		%>
		<%
		if Result ="Modify" then
		%>
		<td><select name="GroupID" class="textfield"><% call SelectGroup1() %></select></td>
        <%
		end if
		%>
	  </tr>
	  <tr>
        <td height="20" align="right">������Ч��</td>
        <td><input name="Working" type="checkbox" value="1" style="HEIGHT: 13px;WIDTH: 13px;" checked></td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;¼&nbsp;����</td>
        <td><input name="AdminName" type="text" class="textfield" id="AdminName" style="WIDTH: 120;" value="<%=AdminName%>" maxlength="16" <%if Result="Modify" then response.write ("readonly")%>>&nbsp;*&nbsp;3-10λ�ַ��������޸�</td>
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
        <td height="20" align="right">����Ա����</td>
        <td><input name="UserName" type="text" class="textfield" id="UserName" style="WIDTH: 120;" value="<%=UserName%>"></td>
      </tr>
      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">����Ȩ�ޣ�</td>
        <td>
		<table border="0">
		<tr>
		
		<td>
	      <input name="Purview309" type="checkbox" value="|309," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|309,")>0 then response.write ("checked")%>> ��վ��Ϣ���ù���
		</td>
		<td style="padding-left:10px">
		 <input name="Purview112" type="checkbox" value="|112," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|112,")>0 then response.write ("checked")%>>&nbsp;��վ��Ϣ����
		</td>
		<td style="padding-left:10px">		
		  <input name="Purview310" type="checkbox" value="|310," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|310,")>0 then response.write ("checked")%>> ������־����
		</td>
		<td style="padding-left:10px">
		<input name="Purview312" type="checkbox" value="|312," style="HEIGHT: 13px;WIDTH: 13px;"
		<%if Instr(Purview,"|312,")>0 then response.write ("checked")%>>  ����վ����
		</td>
		<td style="padding-left:10px">
		<input name="Purview105" type="checkbox" value="|105," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|105,")>0 then response.write ("checked")%>>&nbsp;����Ա���		  
		</td>
		<td style="padding-left:10px">
		   <input name="Purview102" type="checkbox" value="|102," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|102,")>0 then response.write ("checked")%>>&nbsp;����Ա�б�
		 
		</td>
		<td style="padding-left:10px">
		  <input name="Purview101" type="checkbox" value="|101," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|101,")>0 then response.write ("checked")%>>&nbsp;�༭����Ա		  
		</td>
		<td style="padding-left:10px">
		  <input name="Purview111" type="checkbox" checked value="|111," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|111,")>0 then response.write ("checked")%>>&nbsp;�޸�����<font color="red">*</font>
		</td>
		<td style="padding-left:10px">
		   <input name="Purview104" type="checkbox" value="|104," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|104,")>0 then response.write ("checked")%>>&nbsp;��Ա�б�		 
		</td>
		</tr>
		<tr>
		<td>
		  <input name="Purview311" type="checkbox" value="|311," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|311,")>0 then response.write ("checked")%>>  �û�����Ȩ��
          
		</td>
		<td style="padding-left:10px">
		   <input name="Purview103" type="checkbox" value="|103," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|103,")>0 then response.write ("checked")%>>&nbsp;�༭��Ա
			
		</td>
		<td style="padding-left:10px">
		   <input name="Purview307" type="checkbox" value="|307," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|307,")>0 then response.write ("checked")%>>  �޸�����Ȩ��
			
		</td>
		<td style="padding-left:10px">
		   <input name="Purview304" type="checkbox" value="|304," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|304,")>0 then response.write ("checked")%>> ������Ϣɾ��
			
		</td>
		<td style="padding-left:10px">
		    <input name="Purview306" type="checkbox" value="|306," style="HEIGHT: 13px;WIDTH: 13px;"
		    <%if Instr(Purview,"|306,")>0 then response.write ("checked")%>> ������Ϣɾ��
		</td>
		<td style="padding-left:10px">
		   <input name="Purview119" type="checkbox" value="|119," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|119,")>0 then response.write ("checked")%>>&nbsp;�����ƹ�
		</td>
		<td style="padding-left:10px">
		   <input name="Purview118" type="checkbox" value="|118," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|118,")>0 then response.write ("checked")%>>&nbsp;��ʾ����
		</td>
		<td style="padding-left:10px">
		   <input name="Purview120" type="checkbox" value="|120," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|120,")>0 then response.write ("checked")%>>&nbsp;��ʾIP
		</td>
		<td>
		   <input name="Purview120" type="checkbox" value="|121," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|121,")>0 then response.write ("checked")%>>&nbsp;��ʾ�绰����
		</td>
		</tr>
		<tr>
		<td>
		  <input name="Purview316" type="checkbox" value="|316," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|316,")>0 then response.write ("checked")%>> ������Դ	
		</td>
		<td  style="padding-left:10px">
		  <input name="Purview317" type="checkbox" value="|317," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|317,")>0 then response.write ("checked")%>> �������Դ���		
		</td>
		<td  style="padding-left:10px">
		  <input name="Purview300" type="checkbox" value="|300," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|300,")>0 then response.write ("checked")%>> ��ҵ��Ϣ����		
		</td>
		<td style="padding-left:10px">
		  <input name="Purview11" type="checkbox" value="|11," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|11,")>0 then response.write ("checked")%>>&nbsp;�༭��ҵ
		
		</td>
		<td style="padding-left:10px">		    		
		  <input name="Purview12" type="checkbox" value="|12," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|12,")>0 then response.write ("checked")%>>&nbsp;��ҵ�б�
		</td>
		</tr>
		<tr>
		<td>
		   <input name="Purview301" type="checkbox" value="|301," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|301,")>0 then response.write ("checked")%>>  �������Ĺ���
		</td>
		<td style="padding-left:10px">
		  <input name="Purview23" type="checkbox" value="|23," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|23,")>0 then response.write ("checked")%>>&nbsp;�༭����
		</td>		
		<td style="padding-left:10px">
		  <input name="Purview21" type="checkbox" value="|21," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|21,")>0 then response.write ("checked")%>>&nbsp;�������
		</td>
		<td style="padding-left:10px">
		  <input name="Purview22" type="checkbox" value="|22," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|22,")>0 then response.write ("checked")%>>&nbsp;�����б�
		</td>				
		<td style="padding-left:10px">
		  
		</td>
		<td style="padding-left:10px">
		  
		</td>
		</tr>
		<tr>
		<td>
		  <input name="Purview302" type="checkbox" value="|302," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|302,")>0 then response.write ("checked")%>>  ��Ʒչʾ����
		<td style="padding-left:10px">
		  <input name="Purview33" type="checkbox" value="|33," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|33,")>0 then response.write ("checked")%>>&nbsp;�༭��Ʒ
		</td>
		<td style="padding-left:10px">
		  <input name="Purview31" type="checkbox" value="|31," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|31,")>0 then response.write ("checked")%>>&nbsp;��Ʒ���	
		</td>
		<td style="padding-left:10px">
		  <input name="Purview32" type="checkbox" value="|32," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|32,")>0 then response.write ("checked")%>>&nbsp;��Ʒ�б�
		</td>		
		</tr>
		<tr>
		<td>
		   <input name="Purview305" type="checkbox" value="|305," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|305,")>0 then response.write ("checked")%>> ������Ϣ�鿴             
		</td>
		<td style="padding-left:10px">
		   <input name="Purview313" type="checkbox" value="|313," style="HEIGHT: 13px;WIDTH: 13px;"            
		   <%if Instr(Purview,"|313,")>0 then response.write ("checked")%>>  ��������ɾ��
		</td>
		<td style="padding-left:10px">		
		  <input name="Purview314" type="checkbox" value="|314," style="HEIGHT: 13px;WIDTH: 13px;"
          <%if Instr(Purview,"|314,")>0 then response.write ("checked")%>> ����״̬����
		</td>
		<td style="padding-left:10px">
		   <input name="Purview315" type="checkbox" value="|315," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|315,")>0 then response.write ("checked")%>>&nbsp;�����ָ�
		</td>
		<td style="padding-left:10px">
		   <input name="Purview99" type="checkbox" value="|99," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|99,")>0 then response.write ("checked")%>>&nbsp;���������б�
		</td>
		<td style="padding-left:10px">
		  <input name="Purview308" type="checkbox" value="|308," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|308,")>0 then response.write ("checked")%>>  ����������Ϣ����
		</td>
		<td style="padding-left:10px">
		   <input name="Purview93" type="checkbox" value="|93," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|93,")>0 then response.write ("checked")%>>&nbsp;�����б�
		</td>
		<td style="padding-left:10px">
		   <input name="Purview94" type="checkbox" value="|94," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|94,")>0 then response.write ("checked")%>>&nbsp;�����ظ�
		</td>
		<td>
		   <input name="Purview90" type="checkbox" value="|90," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|90,")>0 then response.write ("checked")%>>&nbsp;�������Իظ��鿴
		</td>
		<td></td>
		<td></td>		
		</tr>
		<tr>
		<td>
		   <input name="Purview303" type="checkbox" value="|303," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|303,")>0 then response.write ("checked")%>>   ������Ϣ�鿴�ظ�
		</td>
		<td style="padding-left:10px">
		   <input name="Purview91" type="checkbox" value="|91," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|91,")>0 then response.write ("checked")%>>&nbsp;�����б�
		</td>
		<td style="padding-left:10px">
		   <input name="Purview92" type="checkbox" value="|92," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|92,")>0 then response.write ("checked")%>>&nbsp;�༭����
		</td>
		<td></td>
		</tr>
		</table>
		</td>
      <tr <%if ID<>1 then response.write ("style=display:none")%>>
        <td height="20" align="right">����Ȩ�ޣ�</td>
        <td nowrap><font color="#FF0000">���ó�������Ա�ʺţ������޸ģ�</font></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">��ע˵����</td>
        <td><textarea name="Explain" cols="88" rows="3" class="textfield" id="Explain" style="WIDTH: 580;" ><%=Explain%></textarea></td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="����" style="WIDTH: 60;" ></td>
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
sub AdminEdit()
  dim Action,rsCheckAdd,rs,sql
   dim rspur,sqlpur,leftpur
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭����Ա��Ϣ
    set rs = server.createobject("adodb.recordset")	
    if Result="Add" then '������վ����Ա
		
      set rsCheckAdd = conn.execute("select AdminName from NwebCn_Admin where AdminName='" & trim(Request.Form("AdminName")) & "'")
      if not (rsCheckAdd.bof and rsCheckAdd.eof) then '�жϴ˹���Ա���Ƿ����
        response.write "<script language=javascript> alert('" & trim(Request.Form("AdminName")) & "����Ա�Ѿ����ڣ��뻻һ����¼�������ԣ�');history.back(-1);</script>"
        response.end
      end if  
	  sql="select * from NwebCn_Admin"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("AdminName")))<3 or len(trim(Request.Form("Password")))>10  then
        response.write "<script language=javascript> alert('����Ա��¼��������ַ���Ϊ3-10λ��');history.back(-1);</script>"
        response.end
      end if	  
      if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
        response.write "<script language=javascript> alert('����Ա���������ַ���Ϊ6-16λ��');history.back(-1);</script>"
        response.end
      end if
	  if Request.Form("Password")<>Request.Form("vPassword") then 
        response.write "<script language=javascript> alert('������������벻һ����');history.back(-1);</script>"
        response.end
	  end if
      rs("AdminName")=trim(Request.Form("AdminName"))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
	  GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("GroupName")=GroupIdName(1)
	  rs("Password")=Md5(Request.Form("Password"))
	  rs("UserName")=trim(Request.Form("UserName"))

	  rs("AdminPurview")=Request.Form("Purview11") & Request.Form("Purview12") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") & Request.Form("Purview23") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") &_
	                     Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") &_
	                     Request.Form("Purview71") & Request.Form("Purview72") & Request.Form("Purview73") &_
	                     Request.Form("Purview81") & Request.Form("Purview82") & Request.Form("Purview97") &_
	                     Request.Form("Purview91") & Request.Form("Purview92") & Request.Form("Purview93") &_
	                     Request.Form("Purview94") & Request.Form("Purview95") & Request.Form("Purview96") &_
	                     Request.Form("Purview98") & Request.Form("Purview99") & Request.Form("Purview101") &_
	                     Request.Form("Purview102") & Request.Form("Purview103") & Request.Form("Purview104") &_
	                     Request.Form("Purview105") & Request.Form("Purview111") & Request.Form("Purview112") &_
	                     Request.Form("Purview113") & Request.Form("Purview114") & Request.Form("Purview115") &_
	                     Request.Form("Purview116") & Request.Form("Purview117") & Request.Form("Purview118") &_
	                     Request.Form("Purview119") & Request.Form("Purview120") & Request.Form("Purview300") &_
			     Request.Form("Purview301") & Request.Form("Purview302") & Request.Form("Purview303") &_
	                     Request.Form("Purview304") & Request.Form("Purview305") & Request.Form("Purview90") &_
                             Request.Form("Purview306") & Request.Form("Purview307") & Request.Form("Purview308") &_
	                     Request.Form("Purview309") & Request.Form("Purview310") & Request.Form("Purview311") &_
	                     Request.Form("Purview312") & Request.Form("Purview313")& Request.Form("Purview314")&_
			     Request.Form("Purview315")&Request.Form("Purview316") & Request.Form("Purview317")
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	  
	  
	 
  
   
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Admin where ID="&ID
      rs.open sql,conn,1,3
      rs("AdminName")=trim(Request.Form("AdminName"))
	  'rs("GroupID")=trim(Request.QueryString("GroupID"))
	  'rs("GroupName")=trim(Request.Form("GroupName"))
	  GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("GroupName")=GroupIdName(1)
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
      if trim(Request.Form("Password"))<>"" then
	    if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>20  then
          response.write "<script language=javascript> alert('����Ա���������ַ���Ϊ6-20λ��');history.back(-1);</script>"
          response.end
        end if
	    if Request.Form("Password")<>Request.Form("vPassword") then 
          response.write "<script language=javascript> alert('������������벻һ����');history.back(-1);</script>"
          response.end
	    end if
	    rs("Password")=Md5(Request.Form("Password"))
	  end if
	  rs("UserName")=trim(Request.Form("UserName"))
	  rs("AdminPurview")=Request.Form("Purview11") & Request.Form("Purview12") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") & Request.Form("Purview23") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") &_
	                     Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") & Request.Form("Purview71") &_
	                     Request.Form("Purview72") & Request.Form("Purview73") & Request.Form("Purview81") &_
	                     Request.Form("Purview82") & Request.Form("Purview90") & Request.Form("Purview91") &_
	                     Request.Form("Purview92") & Request.Form("Purview93") & Request.Form("Purview94") &_
	                     Request.Form("Purview95") & Request.Form("Purview96") & Request.Form("Purview97") &_
	                     Request.Form("Purview98") & Request.Form("Purview99") & Request.Form("Purview101") &_
	                     Request.Form("Purview102") & Request.Form("Purview103") & Request.Form("Purview104") &_
	                     Request.Form("Purview105") & Request.Form("Purview111") & Request.Form("Purview112") &_
	                     Request.Form("Purview113") & Request.Form("Purview114") & Request.Form("Purview115") &_
	                     Request.Form("Purview116") & Request.Form("Purview117") & Request.Form("Purview118") &_
	                     Request.Form("Purview119") & Request.Form("Purview120") & Request.Form("Purview300") &_
			     Request.Form("Purview301") & Request.Form("Purview302") & Request.Form("Purview303") &_
	                     Request.Form("Purview304") & Request.Form("Purview305") & Request.Form("Purview306")&_
	                     Request.Form("Purview307") & Request.Form("Purview308") & Request.Form("Purview309") &_
	                     Request.Form("Purview310") & Request.Form("Purview311")& Request.Form("Purview312") &_
			     Request.Form("Purview313") & Request.Form("Purview314")& Request.Form("Purview315") &_
			     Request.Form("Purview316") & Request.Form("Purview317")
	  rs("Explain")=trim(Request.Form("Explain"))
	end if
	rs.update
	rs.close
    set rs=nothing 
	
	 
    response.write "<script language=javascript> alert('�ɹ��༭��վ����Ա��');changeAdminFlag('��վ����Ա');location.replace('AdminList.asp');</script>"
  else '��ȡ����Ա��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Admin where ID="& ID
      rs.open sql,conn,1,1
	  AdminName=rs("AdminName")
	  Working=rs("Working")
	  UserName=rs("UserName")
	  Purview=rs("AdminPurview")
	  Explain=rs("Explain")
	  'GroupID=rs("GroupID")
	  GroupID=rs("GroupID")
	  GroupName=rs("GroupName")
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
  sql="select GroupID,GroupName from NwebCn_MemGroup where GroupID!=1"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("δ�����")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"���橾"&rs("GroupName")&"'")  
    'response.write("<option value='"&rs("GroupID")&"'")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext	
  wend
  rs.close
  set rs=nothing
end sub
%>


<% 
sub SelectGroup1()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from NwebCn_MemGroup where GroupID>1"
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