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
<TITLE>������Ŀ�б�</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|113,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>������Ŀ�����ӣ��޸ĵ�����Ŀ��ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="NavigationEdit.asp?Result=Add" onClick='changeAdminFlag("���ӵ�����Ŀ")'>���ӵ�����Ŀ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NavigationList.asp" onClick='changeAdminFlag("������Ŀ�б�")'>�鿴������Ŀ</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=Navigation" method="post" name="formDel" >
    <tr>
      <td width="18" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
      <td width="28" height="24" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����</font></strong></td>
      <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>��������</strong></font></td>
      <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>���ӵ�ַ</strong></font></td>
      <td width="52" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>��ʾ˳��</strong></font></td>
      <td width="118" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF"><strong>����ʱ��</strong></font></strong></td>
      <td colspan="2" width="76" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="ȫ" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="��" style="HEIGHT: 18px;WIDTH: 16px;">      </td>
    </tr>
	<% NavigationList() %>
  </form>
</table>
<% if request.QueryString("Result")="ModifySequence" then call ModifySequence() %>
<% if request.QueryString("Result")="SaveSequence" then call SaveSequence() %>
</body>
</html>
<%
'-----------------------------------------------------------
function NavigationList()
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
      datafrom="NwebCn_Navigation"
  dim datawhere'��������
      datawhere=""
  dim sqlid'��ҳ��Ҫ�õ���id
  dim Myself,PATH_INFO,QUERY_STRING'��ҳ��ַ�Ͳ���
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'��������
      taxis="order by Sequence asc"
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
    sql="select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)'������ݵ�����
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
      if rs("ViewFlag") then
        Response.Write "<td nowrap><font color='blue'>��</font></td>" & vbCrLf
      else
        Response.Write "<td nowrap><font color='red'>��</font></td>" & vbCrLf
	  end if
      Response.Write "<td nowrap>"&rs("NavName")&"</td>" & vbCrLf
	  if StrLen(rs("NavUrl"))>49 then
        Response.Write "<td title="&rs("NavUrl")&" nowrap>"&StrLeft(rs("NavUrl"),46)&"</td>" & vbCrLf
      else
        Response.Write "<td title="&rs("NavUrl")&" nowrap>"&rs("NavUrl")&"</td>" & vbCrLf
      end if
      Response.Write "<td nowrap><font color='blue'>"&rs("Sequence")&"</font></td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td width='48' nowrap><a href='NavigationEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""�޸ĵ�����Ŀ"")'><font color='#330099'>��</font></a>.<a href='NavigationList.asp?Result=ModifySequence&ID="&rs("ID")&"' onClick='changeAdminFlag(""���򵼺���Ŀ"")'><font color='#330099'>����</font></a></td>" & vbCrLf
 	  Response.Write "<td width='14' nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='6' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='ɾ����ѡ' onClick='ConfirmDel(""�����Ҫɾ����Щ������Ŀ��"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='10' nowrap  bgcolor='#EBF2F9'>���޵�����Ŀ</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='8' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
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
<%
sub ModifySequence()
  dim rs,sql,ID,NavName,Sequence
  ID=request.QueryString("ID")
  set rs = server.createobject("adodb.recordset")
  sql="select * from NwebCn_Navigation where ID="& ID
  rs.open sql,conn,1,1
  NavName=rs("NavName")
  Sequence=rs("Sequence")
  rs.close
  set rs=nothing
  response.write "<br>"
  response.write "<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'>"
  response.write "<form action='NavigationList.asp?Result=SaveSequence' method='post' name='formSequence'>"
  response.write "<tr>"
  response.write "<td height='24' align='center' nowrap  bgcolor='#EBF2F9'>ID��<input name='ID' type='text' class='textfield'  style='WIDTH: 30;' value='"&ID&"' maxlength='4' readonly>&nbsp;������Ŀ���ƣ�<input name='NavName' type='text' class='textfield' id='NavName' style='WIDTH: 180;' value='"&NavName&"' maxlength='36' readonly>&nbsp;����ţ�<input name='Sequence' type='text' class='textfield' style='WIDTH: 60;' value='"&Sequence&"' maxlength='4'  onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('ֻ���������������������');this.value='"&Sequence&"';}"">&nbsp;&nbsp;<input name='submitSequence' type='submit' class='button' value='����' style='WIDTH: 60;' ></td>"
  response.write "</tr>"
  response.write "</form>"
  response.write "</table>"
end sub

sub SaveSequence()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select * from NwebCn_Navigation where ID="& request.form("ID")
  rs.open sql,conn,1,3
  rs("Sequence")=request.form("Sequence")
  rs.update
  rs.close
  set rs=nothing
  response.redirect "NavigationList.asp"
end sub
%>