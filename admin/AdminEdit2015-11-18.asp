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
dim Result
Result=request.QueryString("Result")
dim ID,AdminName,Working,Password,vPassword,UserName,Purview,Explain,AddTime
ID=request.QueryString("ID")
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
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">��&nbsp;¼&nbsp;����</td>
        <td><input name="AdminName" type="text" class="textfield" id="AdminName" style="WIDTH: 120;" value="<%=AdminName%>" maxlength="16" >&nbsp;*&nbsp;3-10λ�ַ��������޸�</td>
      </tr>
      <tr>
        <td height="20" align="right">������Ч��</td>
        <td><input name="Working" type="checkbox" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if Working then response.write ("checked")%>></td>
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
        <td nowrap>
		<input name="Purview309" type="checkbox" value="|309," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|309,")>0 then response.write ("checked")%>> ��վ��Ϣ���ù��� 
		  <input name="Purview112" type="checkbox" value="|112," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|112,")>0 then response.write ("checked")%>>&nbsp;��վ��Ϣ����
		  <input name="Purview11" type="checkbox" value="|11," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|11,")>0 then response.write ("checked")%>>&nbsp;�༭��ҵ
          <input name="Purview12" type="checkbox" value="|12," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|12,")>0 then response.write ("checked")%>>&nbsp;��ҵ�б�
		  <input name="Purview21" type="checkbox" value="|21," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|21,")>0 then response.write ("checked")%>>&nbsp;�������
		  <input name="Purview22" type="checkbox" value="|22," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|22,")>0 then response.write ("checked")%>>&nbsp;�����б�
          <input name="Purview23" type="checkbox" value="|23," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|23,")>0 then response.write ("checked")%>>&nbsp;�༭����
		  <input name="Purview31" type="checkbox" value="|31," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|31,")>0 then response.write ("checked")%>>&nbsp;��Ʒ���
		  <input name="Purview32" type="checkbox" value="|32," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|32,")>0 then response.write ("checked")%>>&nbsp;��Ʒ�б�
          <input name="Purview33" type="checkbox" value="|33," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|33,")>0 then response.write ("checked")%>>&nbsp;�༭��Ʒ
		<input name="Purview300" type="checkbox" value="|300," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|300,")>0 then response.write ("checked")%>> ��ҵ��Ϣ���� 
          
           <input name="Purview301" type="checkbox" value="|301," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|301,")>0 then response.write ("checked")%>>  �������Ĺ���
          
           <input name="Purview302" type="checkbox" value="|302," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|302,")>0 then response.write ("checked")%>>  ��Ʒչʾ���� </td>
      </tr>
      <!--<tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td>
		  <input name="Purview41" type="checkbox" value="|41," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|41,")>0 then response.write ("checked")%>>&nbsp;�������
		  <input name="Purview42" type="checkbox" value="|42," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|42,")>0 then response.write ("checked")%>>&nbsp;�����б�
          <input name="Purview43" type="checkbox" value="|43," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|43,")>0 then response.write ("checked")%>>&nbsp;�༭����
		  <input name="Purview51" type="checkbox" value="|51," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|51,")>0 then response.write ("checked")%>>&nbsp;�������
		  <input name="Purview52" type="checkbox" value="|52," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|52,")>0 then response.write ("checked")%>>&nbsp;�����б�
          <input name="Purview53" type="checkbox" value="|53," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|53,")>0 then response.write ("checked")%>>&nbsp;�༭����
		  <input name="Purview61" type="checkbox" value="|61," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|61,")>0 then response.write ("checked")%>>&nbsp;��Ƹ�б�
		  <input name="Purview62" type="checkbox" value="|62," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|62,")>0 then response.write ("checked")%>>&nbsp;�༭��Ƹ</td>
      </tr>-->
      <!--<tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td>
		  <input name="Purview71" type="checkbox" value="|71," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|71,")>0 then response.write ("checked")%>>&nbsp;�������
		  <input name="Purview72" type="checkbox" value="|72," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|72,")>0 then response.write ("checked")%>>&nbsp;�����б�
          <input name="Purview73" type="checkbox" value="|73," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|73,")>0 then response.write ("checked")%>>&nbsp;�༭����
		  <input name="Purview81" type="checkbox" value="|81," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|81,")>0 then response.write ("checked")%>>&nbsp;����б�
		  <input name="Purview82" type="checkbox" value="|82," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|82,")>0 then response.write ("checked")%>>&nbsp;�༭���
	</td>
      </tr>-->
      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td>
		  <!--<input name="Purview95" type="checkbox" value="|95," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|95,")>0 then response.write ("checked")%>>&nbsp;��Ӧ�б�
		  <input name="Purview96" type="checkbox" value="|96," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|96,")>0 then response.write ("checked")%>>&nbsp;��Ӧ�ظ�
		  <input name="Purview97" type="checkbox" value="|97," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|97,")>0 then response.write ("checked")%>>&nbsp;�˲��б�
		  <input name="Purview98" type="checkbox" value="|98," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|98,")>0 then response.write ("checked")%>>&nbsp;�ظ��˲�-->		  		  
		  <input name="Purview101" type="checkbox" value="|101," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|101,")>0 then response.write ("checked")%>>&nbsp;�༭����Ա
		  <input name="Purview102" type="checkbox" value="|102," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|102,")>0 then response.write ("checked")%>>&nbsp;����Ա�б�
          <input name="Purview103" type="checkbox" value="|103," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|103,")>0 then response.write ("checked")%>>&nbsp;�༭��Ա
		  <input name="Purview118" type="checkbox" value="|118," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|118,")>0 then response.write ("checked")%>>&nbsp;��ʾ����
		  <input name="Purview120" type="checkbox" value="|120," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|120,")>0 then response.write ("checked")%>>&nbsp;��ʾIP
		<input name="Purview304" type="checkbox" value="|304," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|304,")>0 then response.write ("checked")%>> ������Ϣɾ��
          <input name="Purview311" type="checkbox" value="|311," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|311,")>0 then response.write ("checked")%>>  �û�����Ȩ��
         <input name="Purview312" type="checkbox" value="|312," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|312,")>0 then response.write ("checked")%>>  ����վ����
            <input name="Purview313" type="checkbox" value="|313," style="HEIGHT: 13px;WIDTH: 13px;"
            
		  <%if Instr(Purview,"|313,")>0 then response.write ("checked")%>>  ��������ɾ��</td>
      </tr>
      <!--<tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">&nbsp;</td>
        <td>
		  <input name="Purview113" type="checkbox" value="|113," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|113,")>0 then response.write ("checked")%>>&nbsp;������Ŀ		  		  
		  <input name="Purview114" type="checkbox" value="|114," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|114,")>0 then response.write ("checked")%>>&nbsp;��������
		  <input name="Purview115" type="checkbox" value="|115," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|115,")>0 then response.write ("checked")%>>&nbsp;���ݿ����
          <input name="Purview116" type="checkbox" value="|116," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|116,")>0 then response.write ("checked")%>>&nbsp;�ռ�ͳ��
		  <input name="Purview117" type="checkbox" value="|117," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|117,")>0 then response.write ("checked")%>>&nbsp;����ͳ�Ʋ鿴
		  <input name="Purview119" type="checkbox" value="|119," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|119,")>0 then response.write ("checked")%>>&nbsp;��������
          </td>
      </tr>-->

      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" rowspan="2" align="right">&nbsp;</td>
        <td>

          </td>
      </tr>

      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td>
		  <input name="Purview91" type="checkbox" value="|91," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|91,")>0 then response.write ("checked")%>>&nbsp;�����б�<font color="red">*</font>

          <input name="Purview92" type="checkbox" value="|92," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|92,")>0 then response.write ("checked")%>>&nbsp;�༭����<font color="red">*</font>
          <input name="Purview93" type="checkbox" value="|93," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|93,")>0 then response.write ("checked")%>>&nbsp;�����б�<font color="red">*</font>
		  <input name="Purview94" type="checkbox" value="|94," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|94,")>0 then response.write ("checked")%>>&nbsp;�����ظ�<font color="red">*</font>
		  <input name="Purview104" type="checkbox" value="|104," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|104,")>0 then response.write ("checked")%>>&nbsp;��Ա�б�
		  <input name="Purview105" type="checkbox" value="|105," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|105,")>0 then response.write ("checked")%>>&nbsp;��Ա���
		  <input name="Purview111" type="checkbox" value="|111," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|111,")>0 then response.write ("checked")%>>&nbsp;�޸�����<font color="red">*</font>          
          <input name="Purview303" type="checkbox" value="|303," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|303,")>0 then response.write ("checked")%>>   ������Ϣ�鿴�ظ�<font color="red">*</font> 
          
          <input name="Purview305" type="checkbox" value="|305," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|305,")>0 then response.write ("checked")%>> ������Ϣ�鿴<font color="red">*</font>
          
           <input name="Purview306" type="checkbox" value="|306," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|306,")>0 then response.write ("checked")%>> ������Ϣɾ��<font color="red">*</font>
          
          
           <input name="Purview307" type="checkbox" value="|307," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|307,")>0 then response.write ("checked")%>>  �޸�����Ȩ��<font color="red">*</font>
           <input name="Purview314" type="checkbox" value="|314," style="HEIGHT: 13px;WIDTH: 13px;"
            <%if Instr(Purview,"|314,")>0 then response.write ("checked")%>> ����״̬����<font color="red">*</font>
            <input name="Purview308" type="checkbox" value="|308," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|308,")>0 then response.write ("checked")%>>  ����������Ϣ����<font color="red">*</font>
          
          <input name="Purview310" type="checkbox" value="|310," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|310,")>0 then response.write ("checked")%>> ������־����<font color="red">*</font>
           </td>
      </tr>
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
	                     Request.Form("Purview98") & Request.Form("Purview98") & Request.Form("Purview101") &_
	                     Request.Form("Purview102") & Request.Form("Purview103") & Request.Form("Purview104") &_
	                     Request.Form("Purview105") & Request.Form("Purview111") & Request.Form("Purview112") &_
	                     Request.Form("Purview113") & Request.Form("Purview114") & Request.Form("Purview115") &_
	                     Request.Form("Purview116") & Request.Form("Purview117") & Request.Form("Purview118") &_
	                     Request.Form("Purview119") & Request.Form("Purview120") & Request.Form("Purview300") &_
			     Request.Form("Purview301") & Request.Form("Purview302") & Request.Form("Purview303") &_
	                     Request.Form("Purview304") & Request.Form("Purview305") &_
                             Request.Form("Purview306") &_
	                     Request.Form("Purview307") & Request.Form("Purview308") & Request.Form("Purview309") &_
	                     Request.Form("Purview310") & Request.Form("Purview311")& Request.Form("Purview312")&_
			     Request.Form("Purview313")&Request.Form("Purview314")
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	  
	  
	 
  
   
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_Admin where ID="&ID
      rs.open sql,conn,1,3
      rs("AdminName")=trim(Request.Form("AdminName"))
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
	                     Request.Form("Purview61") & Request.Form("Purview62") &_
	                     Request.Form("Purview71") & Request.Form("Purview72") & Request.Form("Purview73") &_
	                     Request.Form("Purview81") & Request.Form("Purview82") & Request.Form("Purview97") &_
	                     Request.Form("Purview91") & Request.Form("Purview92") & Request.Form("Purview93") &_
	                     Request.Form("Purview94") & Request.Form("Purview95") & Request.Form("Purview96") &_
	                     Request.Form("Purview98") & Request.Form("Purview98") & Request.Form("Purview101") &_
	                     Request.Form("Purview102") & Request.Form("Purview103") & Request.Form("Purview104") &_
	                     Request.Form("Purview105") & Request.Form("Purview111") & Request.Form("Purview112") &_
	                     Request.Form("Purview113") & Request.Form("Purview114") & Request.Form("Purview115") &_
	                     Request.Form("Purview116") & Request.Form("Purview117") & Request.Form("Purview118") &_
	                     Request.Form("Purview119") & Request.Form("Purview120") & Request.Form("Purview300") &_
			     Request.Form("Purview301") & Request.Form("Purview302") & Request.Form("Purview303") &_
	                     Request.Form("Purview304") & Request.Form("Purview305") & Request.Form("Purview306")&_
	                     Request.Form("Purview307") & Request.Form("Purview308") & Request.Form("Purview309") &_
	                     Request.Form("Purview310") & Request.Form("Purview311")& Request.Form("Purview312")&_
						 Request.Form("Purview313") & Request.Form("Purview314")
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
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
  
%>
