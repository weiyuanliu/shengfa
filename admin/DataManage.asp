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
<TITLE>���ݿ����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|310,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>���ݿ������ϵͳ���ݱ��֣�ѹ�����ָ�������Ա��¼��־</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9">
    <a href="DataManage.asp" onClick='changeAdminFlag("���ݿ����")'>��Ŀ��ҳ</a><font color="#0000FF">&nbsp;|&nbsp;</font>��վ���ݿ⣺<a href="DataManage.asp?Action=DataBackup&Result=Site" onClick='changeAdminFlag("��վ���ݿⱸ��")'>����</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataCompact&Result=Site" onClick='changeAdminFlag("��վѹ�����ݿ�")'>ѹ��</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataResume&Result=Site" onClick='changeAdminFlag("��վ�ָ����ݿ�")'>�ָ�</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DataManage.asp?Action=DataLog" onClick='changeAdminFlag("����Ա��¼��־")'>����Ա��¼��־</a><font color="#0000FF">&nbsp;|&nbsp;</font>�������ݿ⣺<a href="DataManage.asp?Action=DataBackup&Result=Stat" onClick='changeAdminFlag("�������ݿⱸ��")'>����</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataCompact&Result=Stat" onClick='changeAdminFlag("����ѹ�����ݿ�")'>ѹ��</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataResume&Result=Stat" onClick='changeAdminFlag("�����ָ����ݿ�")'>�ָ�</a><font color="#0000FF">&nbsp;</font></td>    
  </tr>
</table>
<br>
<% call DataManage() %>
</body>
</html>
<%
sub DataManage()
  Dim Action
  Action=request.QueryString("Action")
  Select Case Action
    Case "DataBackup"
	  DataBackup
    Case "DataCompact"
	  DataCompact
    Case "DataResume"
	  DataResume
    Case "DataLog"
	  DataLog
    Case Else
      DataMain
  End Select
end sub  
%>

<%
function DataMain
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9'>")
  response.write ("����˵����<br>���������ݿ��������Ϊ[����&nbsp;��&nbsp;ѹ��&nbsp;��&nbsp;�ָ�]<br>����������ǰ�����[<font color='#330099'>����</font>]���ݿ⣬����ʹ���е����ݿⲻ�ܱ�ѹ��<BR>�������ָ����ݿ�ʱ���Ḳ�ǵ�ǰʹ���е����ݿ�<br>����������Ա��¼��־�����鿴��ɾ��")
  response.write ("</td></tr></table>")
end function

function DataBackup()
  dim From,Fso,Result
  From=request.QueryString("From")
  Result=request.QueryString("Result")
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9' align='center'>")
  response.write ("<table width='560' border='0' cellspacing='0' cellpadding='0'><tr><td height='16'></td></tr>")
  response.write ("<tr><td height='20'>˵�����޸����ݿⱸ�ݱ���·�����ļ����������[ϵͳ���á�վ�㳣�����á����ݿⱸ��·��]</td></tr>")
  if From="Confirm" then
    set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if Result="Site" then
	  Fso.CopyFile Server.MapPath(SiteDataPath),Server.MapPath(SiteDataBakPath)
      response.write ("<tr><td height='20'>�ɹ������Ѿ��ɹ��������ݵ�&nbsp;<a href='"&SiteDataBakPath&"' target='_blank'><font color='#330099'>"&SiteDataBakPath&"</font></a>&nbsp;��ע�⼰ʱɾ�����õı��ݣ�</td></tr>")
	else
	  Fso.CopyFile Server.MapPath(StatDataPath),Server.MapPath(StatDataBakPath)
      response.write ("<tr><td height='20'>�ɹ������Ѿ��ɹ��������ݵ�&nbsp;<a href='"&StatDataBakPath&"' target='_blank'><font color='#330099'>"&StatDataBakPath&"</font></a>&nbsp;��ע�⼰ʱɾ�����õı��ݣ�</td></tr>")
    end if
 	response.write ("<tr><td height='20'>�汾�����ݿ��ʱ��汾Ϊ&nbsp;"& now() &"</td></tr>")
    Set Fso=nothing
  end if	  
  response.write ("<form id='DataBackupForm' name='DataBackupForm' method='post' action='DataManage.asp?From=Confirm&Action=DataBackup&Result="&Result&"'>")
  if Result="Site" then
    response.write ("<tr><td height='30'>��Դ��<input name='fromPath' readonly type='text' size='60' value='"&SiteDataPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>Ŀ�꣺<input name='toPath' readonly type='text' size='60' value='"&SiteDataBakPath&"' class='textfield' /></td></tr>")
  else
    response.write ("<tr><td height='30'>��Դ��<input name='fromPath' readonly type='text' size='60' value='"&StatDataPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>Ŀ�꣺<input name='toPath' readonly type='text' size='60' value='"&StatDataBakPath&"' class='textfield' /></td></tr>")
  end if
  response.write ("<tr><td height='30'><input type='submit' value='ȷ������' class='button' /></td></tr>")
  response.write ("</form>")  
  response.write ("<tr><td height='16'></td></tr></table>")
  response.write ("</td></tr></table>")
end function

function DataCompact()
  dim From,Fso,Engine,SDBPath,Result
  From=request.QueryString("From")
  Result=request.QueryString("Result")
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9' align='center'>")
  response.write ("<table width='560' border='0' cellspacing='0' cellpadding='0'><tr><td height='16'></td></tr>")
  response.write ("<tr><td height='20'>˵����ѹ��ǰ�����[<font color='#330099'>����</font>]���ݿ⣬����ʹ���е����ݿⲻ�ܱ�ѹ��</td></tr>")
  if From="Confirm" then
    if Result="Site" then
      SDBPath = server.mappath(SiteDataBakPath)
	else
      SDBPath = server.mappath(StatDataBakPath)
	end if
    set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if Fso.FileExists(SDBPath) then
      Set Engine =Server.CreateObject("JRO.JetEngine")
	  if request("boolIs") = "97" then
	     Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath, _
		                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath & "_temp.mdb;" _
		                        & "Jet OLEDB:Engine Type=" & JET_3X
	  else 
	     Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath, _
		                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath & "_temp.mdb"
      end if
      Fso.CopyFile SDBPath & "_temp.mdb",SDBPath
      Fso.DeleteFile(SDBPath & "_temp.mdb")
      set Fso = nothing
      set Engine = nothing
      response.write ("<tr><td height='20'>�ɹ������ݿ�&nbsp;<a href='"&SDBPath&"' target='_blank'><font color='#330099'>"&SiteDataBakPath&"</font></a>&nbsp;�Ѿ�ѹ���ɹ���</td></tr>")
	  response.write ("<tr><td height='20'>�汾�����ݿ��ʱ��汾Ϊ&nbsp;"& now() &"</td></tr>")
    else
      response.write ("<tr><td height='20'>ʧ�ܣ����ݿ�&nbsp;<a href='"&SDBPath&"' target='_blank'><font color='#330099'>"&SiteDataBakPath&"</font></a>&nbsp;ѹ��ʧ�ܣ�����·�������ݿ����Ƿ���ڣ�</td></tr>")
    end if
  end if
  response.write ("<form id='DataCompactForm' name='DataCompactForm' method='post' action='DataManage.asp?From=Confirm&Action=DataCompact&Result="&Result&"'>")
  if Result="Site" then
    response.write ("<tr><td height='30'>Ŀ�꣺<input name='toPath' readonly type='text' size='60' value='"&SiteDataBakPath&"' class='textfield'/></td></tr>")
  else
    response.write ("<tr><td height='30'>Ŀ�꣺<input name='toPath' readonly type='text' size='60' value='"&StatDataBakPath&"' class='textfield'/></td></tr>")
  end if
  response.write ("<tr><td height='30'><input type='submit' value='ȷ��ѹ��' class='button' /></td></tr>")
  response.write ("</form>")  
  response.write ("<tr><td height='16'></td></tr></table>")
  response.write ("</td></tr></table>")
end function

function DataResume()
  dim From,Fso,SDPath,SDBPath,Result
  From=request.QueryString("From")
  Result=request.QueryString("Result")
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9' align='center'>")
  response.write ("<table width='560' border='0' cellspacing='0' cellpadding='0'><tr><td height='16'></td></tr>")
  response.write ("<tr><td height='20'>˵�����޸ı��ݡ�Ŀ�����ݿ�ı���·�����ļ����������[ϵͳ���á�վ�㳣�����á����ݿⱸ��·��]</td></tr>")
  if From="Confirm" then
    if Result="Site" then
	  SDPath = server.mappath(SiteDataPath)
      SDBPath = server.mappath(SiteDataBakPath)
	else
	  SDPath = server.mappath(StatDataPath)
      SDBPath = server.mappath(StatDataBakPath)
	end if
    set Fso=Server.CreateObject("Scripting.FileSystemObject")
    if Fso.FileExists(SDBPath) then
      Fso.CopyFile SDBPath,SDPath
      Set Fso=nothing
      response.write ("<tr><td height='20'>�ɹ������Ѿ��ɹ��ָ����ݿ�&nbsp;<font color='#330099'>"&SDPath&"</font>&nbsp;ע�⼰ʱɾ�����õı��ݣ�</td></tr>")
	  response.write ("<tr><td height='20'>�汾�����ݿ��ʱ��汾Ϊ&nbsp;"& now() &"</td></tr>")
    else
      response.write ("<tr><td height='20'>ʧ�ܣ����ݿ�&nbsp;<a href='"&SDBPath&"' target='_blank'><font color='#330099'>"&SDBPath&"</font></a>&nbsp;ѹ��ʧ�ܣ�����·�������ݿ����Ƿ���ڣ�</td></tr>")
    end if
  end if	    
  response.write ("<form id='DataResumeForm' name='DataResumeForm' method='post' action='DataManage.asp?From=Confirm&Action=DataResume&Result="&Result&"'>")
  if  Result="Site" then
    response.write ("<tr><td height='30'>��Դ��<input name='fromPath' readonly type='text' size='60' value='"&SiteDataBakPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>Ŀ�꣺<input name='toPath' readonly type='text' size='60' value='"&SiteDataPath&"' class='textfield' /></td></tr>")
  else
    response.write ("<tr><td height='30'>��Դ��<input name='fromPath' readonly type='text' size='60' value='"&StatDataBakPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>Ŀ�꣺<input name='toPath' readonly type='text' size='60' value='"&StatDataPath&"' class='textfield' /></td></tr>")
  end if
  response.write ("<tr><td height='30'><input type='submit' value='ȷ���ָ�' class='button' /></td></tr>")
  response.write ("</form>")  
  response.write ("<tr><td height='16'></td></tr></table>")
  response.write ("</td></tr></table>")
end function

function DataLog()
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=LoginLog" method="post" name="formDel" >
    <tr>
      <td width="60" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
      <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>��¼��</strong></font></td>
      <td width="70" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>�û���</strong></font></td>
      <td width="124" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong><font color="#FFFFFF">��¼IP</font></strong></font></td>
      <td width="260" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">��¼ʱ�����</font></strong></td>
      <td width="124" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF"><strong>����ʱ��</strong></font></strong></td>
      <td nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="ȫ" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="��" style="HEIGHT: 18px;WIDTH: 16px;">
	  </td>
    </tr>
	<% AdminLoginLog() %>
  </form>
</table>
<%
end function
%>
<%
'-----------------------------------------------------------
function AdminLoginLog()
  dim idCount'��¼����
  dim pages'ÿҳ����
      pages=20
  dim pagec'��ҳ��
  dim page'ҳ��
      page=clng(request("Page"))
  dim pagenc 'ÿҳ��ʾ�ķ�ҳҳ������=pagenc*2+1
      pagenc=2
  dim pagenmax 'ÿҳ��ʾ�ķ�ҳ�����ҳ��
  dim pagenmin 'ÿҳ��ʾ�ķ�ҳ����Сҳ��
  dim datafrom'���ݱ���
      datafrom="NwebCn_AdminLog"
  dim datawhere'��������
      datawhere=""
  dim sqlid'��ҳ��Ҫ�õ���id
  dim Myself,PATH_INFO,QUERY_STRING'��ҳ��ַ�Ͳ���
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" then
	    Myself = PATH_INFO & "?"
	  else
	  	if Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
          Myself = PATH_INFO & "?" & QUERY_STRING & "&"
		else
	      Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
		end if
	  end if
  dim taxis'�������� asc, desc
      taxis="order by id desc"
  dim i'����ѭ��������
  dim rs,sql'sql���
  '��ȡ��¼����
  sql="select count(ID) as idCount from ["& datafrom &"]" & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,0,1
  idCount=rs("idCount")
  '��ȡ��¼����

  if(idcount>0) then'�����¼����=0,�򲻴���
    if(idcount mod pages=0)then'�����¼��������ÿҳ����������,��=��¼����/ÿҳ����+1
	  pagec=int(idcount/pages)'��ȡ��ҳ��
   	else
      pagec=int(idcount/pages)+1'��ȡ��ҳ��
    end if
	'��ȡ��ҳ��Ҫ�õ���id============================================
    '��ȡ���м�¼��id��ֵ,��Ϊֻ��id�����ٶȺܿ�
    sql="select id from ["& datafrom &"] " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    rs.pagesize = pages 'ÿҳ��ʾ��¼��
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("id")
	  else
	    sqlid=sqlid &","&rs("id")
	  end if
	  rs.movenext
    next
  '��ȡ��ҳ��Ҫ�õ���id����============================================
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  if(idcount>0 and sqlid<>"") then'�����¼����=0,�򲻴���
    '��inˢѡ��ҳ�����Ե�����,����ȡ��ҳ���������,�����ٶȿ�
    sql="select [ID],[AdminName],[UserName],[LoginIP],[LoginSoft],[LoginTime] from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)'������ݵ����
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("AdminName")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("UserName")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("LoginIP")&"</td>" & vbCrLf
	  if len(rs("LoginSoft"))>40 then
        Response.Write "<td nowrap title='�������&#13;"&rs("LoginSoft")&"'>"&left(rs("LoginSoft"),40)&"...</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='�������&#13;"&rs("LoginSoft")&"'>"&rs("LoginSoft")&"</td>" & vbCrLf
      end if 
      Response.Write "<td nowrap>"&rs("LoginTime")&"</td>" & vbCrLf
 	  Response.Write "<td nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='6' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='ɾ����ѡ'  onClick='ConfirmDel(""�����Ҫɾ����Щ����Ա��¼��־��"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write ("<tr><td height='50' align='center' colspan='7' nowrap  bgcolor='#EBF2F9'>���޹���Ա��¼��־</td></tr>")
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='7' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td>���ƣ�<font color='#ff6600'>"&idcount&"</font>����¼&nbsp;ҳ�Σ�<font color='#ff6600'>"&page&"</font></strong>/"&pagec&"&nbsp;ÿҳ��<font color='#ff6600'>"&pages&"</font>��</td>" & vbCrLf
  Response.Write "<td align='right'>" & vbCrLf
  '���÷�ҳҳ�뿪ʼ===============================
  pagenmin=page-pagenc '����ҳ�뿪ʼֵ
  pagenmax=page+pagenc '����ҳ�����ֵ
  if(pagenmin<1) then pagenmin=1 '���ҳ�뿪ʼֵС��1��=1
  if(page>1) then response.write ("<a href='"& myself &"Page=1'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>9</font></a>&nbsp;") '���ҳ�����1����ʾ(��һҳ)
  if(pagenmin>1) then response.write ("<a href='"& myself &"Page="& page-(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>7</font></a>&nbsp;") '���ҳ�뿪ʼֵ����1����ʾ(��ǰ)
  if(pagenmax>pagec) then pagenmax=pagec '���ҳ�����ֵ������ҳ��,��=��ҳ��
  for i = pagenmin to pagenmax'ѭ�����ҳ��
	if(i=page) then
	  response.write ("&nbsp;<font color='#ff6600'>"& i &"</font>&nbsp;")
	else
	  response.write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write ("&nbsp;<a href='"& myself &"Page="& page+(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>8</font></a>&nbsp;") '���ҳ�����ֵС����ҳ������ʾ(����)
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>:</font></a>&nbsp;") '���ҳ��С����ҳ������ʾ(���ҳ)	
  '���÷�ҳҳ�����===============================
  Response.Write "��������&nbsp;<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('ֻ������תĿ��ҳ��������������');this.value='"&Page&"';}"" style='HEIGHT: 18px;WIDTH: 40px;'  type='text' class='textfield' value='"&Page&"'>&nbsp;ҳ" & vbCrLf
  Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick='GoPage("""&Myself&""")' value='GO'>" & vbCrLf
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf
  rs.close
  set rs=nothing
  Response.Write "</td>" & vbCrLf  
  Response.Write "</tr>" & vbCrLf
'-----------------------------------------------------------
'-----------------------------------------------------------
end function 
%>