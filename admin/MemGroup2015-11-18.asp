<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'����������������������������������������������������������������
'����������������������������������������������������������������
'�������������������տƼ���ҵ��վ����ϵͳ���Σףţ£�����������������
'����������������������������������������������������������������
'������Ȩ���С�lisuo.com
'
'�����������������տƼ�������
'��������������Add:�Ĵ�ʡ�����������228��/611930
'��������������Tel:028-68067902  Fax:83708850
'��������������E-m:duolaimi-123@163.com
'��������������Q Q:59309100
'
'���������ַ��[��Ʒ����]http://www.qisehu.com
'��������������[֧����̳]http://www.qisehu.com/bbs
'
'������ʾ��ַ��http://www.qisehu.com
'����������������������������������������������������������������
'����������������������������������������������������������������
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>����Ա���</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="�ɶ����տƼ����޹�˾,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|105,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��վ��Ա������Ա������ã���ӣ��޸Ļ�Ա��Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="MemGroup.asp?Result=Add" onClick='changeAdminFlag("��ӹ���Ա���")'>��ӹ���Ա���</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="AdminList.asp" onClick='changeAdminFlag("�鿴���й���Ա")'>�鿴���й���Ա</a></td>    
  </tr>
</table>
<br>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,GroupID,GroupName,GroupLevel,Explain,AddTime,RanNum
ID=request.QueryString("ID")
randomize timer
RanNum=Int((8999)*Rnd +1009)
if Result<>"" then
  call MemGroupEdit() 
end if
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=MemGroup" method="post" name="formDel" >
    <tr>
      <td width="30" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
      <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>����Ա����</strong></font></td>
      <td width="68" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>�������</strong></font></td>
      <td nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">˵��</font></strong></td>
      <td width="118" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF"><strong>����ʱ��</strong></font></strong></td>
      <td width="76" colspan="2" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="ȫ" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="��" style="HEIGHT: 18px;WIDTH: 16px;">      </td>
    </tr>
	<% MemGroupList() %>
  </form>
</table>
</BODY>
</HTML>
<%
sub MemGroupEdit()
  dim Action,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭��Ա�����Ϣ
    set rs = server.createobject("adodb.recordset")
    if Result="Add" then '������Ա���
	  sql="select * from NwebCn_MemGroup"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language=javascript> alert('��Ա������Ʊ�����ַ���Ϊ6-16�ַ���3-8�����֣�');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupID")=Request.Form("GroupID")
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸���վ����Ա
      sql="select * from NwebCn_MemGroup where ID="&ID
      rs.open sql,conn,1,3
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language=javascript> alert('����Ա������Ʊ�����ַ���Ϊ6-16�ַ���3-8�����֣�');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("Explain")=trim(Request.Form("Explain"))
      conn.execute("Update NwebCn_Members set GroupName='"&trim(Request.Form("GroupName"))&"' where GroupID='"&trim(Request.Form("GroupID"))&"'")
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭����Ա���');changeAdminFlag('����ԱԱ���');location.replace('MemGroup.asp');</script>"
  else '��ȡ����Ա��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_MemGroup where ID="& ID
      rs.open sql,conn,1,1
	  if rs.RecordCount=0 then
        response.write "<script language=javascript> alert('���ݿ����޴˼�¼����ȷ�����أ�');history.back(-1)</script>"
        response.end
	  end if
	  ID=rs("ID")
      GroupID=rs("GroupID")
	  GroupName=rs("GroupName")
	  Explain=rs("Explain")
	  rs.close
      set rs=nothing 
	end if
  end if
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editMemGroup" method="post" action="MemGroup.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onSubmit="return CheckMemGroup()">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">I D��</td>
        <td><input name="ID" type="text" class="textfield" id="ID" style="WIDTH: 100;" value="<%if ID="" then response.write ("�Զ�") else response.write (ID) end if%>" maxlength="6" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">����Ա��ţ�</td>
        <td><input name="GroupID" type="text" class="textfield" id="GroupID" style="WIDTH: 100;" value="<%=GroupID%>" maxlength="4" >&nbsp;*���ű���Ϊ���֣�����Ψһ</td>
      </tr>
      <tr>
        <td height="20" align="right">������ƣ�</td>
        <td><input name="GroupName" type="text" class="textfield" id="GroupName" style="WIDTH: 100;" value="<%=GroupName%>">&nbsp;*����Ա������Ʊ�����ַ���Ϊ6-16�ַ���3-8�����֣�</td>
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
<br>
<%
end sub
'-----------------------------------------------------------
function MemGroupList()
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
      datafrom="NwebCn_MemGroup"
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
      taxis="order by id asc"
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
    while(not rs.eof)'������ݵ����
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("GroupID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("GroupName")&"</td>" & vbCrLf
	  if len(rs("Explain"))>24 then
        Response.Write "<td nowrap title='˵����&#13;"&rs("Explain")&"'>"&left(rs("Explain"),22)&"...</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='˵����&#13;"&rs("Explain")&"'>"&rs("Explain")&"</td>" & vbCrLf
      end if 
      Response.Write "<td nowrap>"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td width='40' nowrap><a href='MemGroup.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""�޸Ĺ���Ա���"")'><font color='#330099'>�޸�</font></a></td>" & vbCrLf
      if rs("ID")=1 then
	    Response.Write "<td width='22' nowrap></td>" & vbCrLf
      else
 	    Response.Write "<td width='22' nowrap><input name='selectID' type='checkbox' value='"&rs("GroupID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      end if
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='5' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td nowrap colspan='2'  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='ɾ����ѡ'  onClick='ConfirmDel(""�����Ҫɾ����Щ����Ա�����"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write ("<tr><td height='50' align='center' colspan='8' nowrap  bgcolor='#EBF2F9'>���޻�Ա���</td></tr>")
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