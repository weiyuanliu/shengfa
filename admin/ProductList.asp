<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'����������������������������������������������������������������
'����������������������������������������������������������������
'��������������������ҵ��վ����ϵͳ���Σףţ£�����������������
'����������������������������������������������������������������
'������Ȩ���С�Nweb.cn
'
'���������������򲩹�����
'��������������Add:�Ĵ�ʡ�����������228��/611930
'��������������Tel:028-88079088  Fax:83708850
'��������������E-m:duolaimi-123@163.com
'��������������Q Q:59309100
'
'���������ַ��[��Ʒ����]http://www.Nweb.cn
'��������������[֧����̳]http://www.Nweb.cn/bbs
'
'������ʾ��ַ��http://www.Nweb.cn
'����������������������������������������������������������������
'����������������������������������������������������������������
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - NWEB.CN-STUDIO" />
<META NAME="Author" CONTENT="�����缼��������,www.Nweb.cn" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>��Ʒ�б�</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|32,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<%
dim Result,StartDate,EndDate,Keyword,SortID,SortPath
Result=request.QueryString("Result")
StartDate=request.QueryString("StartDate")
EndDate=request.QueryString("EndDate")
Keyword=request.QueryString("Keyword")
SortID=request.QueryString("SortID")
SortPath=request.QueryString("SortPath")
function PlaceFlag()
  if Result="Search" then
    Response.Write "��Ʒ���б�&nbsp;->&nbsp;����&nbsp;->&nbsp;���ʱ��[<font color='red'>"&StartDate&"��"&EndDate&"</font>]���ؼ���[<font color='red'>"&Keyword&"</font>]"
  else
    if SortPath<>"" then
      Response.Write "��Ʒ���б�&nbsp;->&nbsp<a href='ProductList.asp'>ȫ��</a>"
	  TextPath(SortID)
	else
      Response.Write "��Ʒ���б�&nbsp;->&nbspȫ��"
	end if
  end if
end function  
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��Ʒ����������鿴����ӣ��޸ģ�ɾ����Ʒ��Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="Search.asp?Result=Products">
          <td nowrap> ��Ʒ��������
          <script language=javascript> 
          var myDate=new dateSelector(); 
          myDate.year--; 
		  myDate.date; 
          myDate.inputName='start_date';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name�� 
          myDate.display(); 
          </script>
          &nbsp;��
          <script language=javascript> 
          myDate.year++; 
          myDate.inputName='end_date';  //ע����������������name��ͬһҳ�е���������򣬲��ܳ����ظ���name�� 
          myDate.display(); 
          </script>
          &nbsp;&nbsp;�ؼ��֣�<input name="Keyword" type="text" class="textfield" value="<%=Keyword%>" size="18">
          <input name="submitSearch" type="submit" class="button" value="����">
          </td>
        </form>
        <td align="right" nowrap>�鿴��<a href="ProductList.asp" onClick='changeAdminFlag("��Ʒ�б�")'>ȫ����Ʒ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="ProductSort.asp" onClick='changeAdminFlag("ѡ��鿴����")'>��������Ʒ</a></td>
      </tr>
    </table>      </td>    
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"><%PlaceFlag()%></td>
  </tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=Products" method="post" name="formDel" >
  <tr>
    <td width="30" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
    <td width="28" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>����</strong></font></td>
    <td width="90" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>��Ʒ���</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>��Ʒ����</strong>��ָ����ʾ�������</font></td>
    <td width="30" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>״̬</strong></font></td>
    <td width="80" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">��ʾ˳��</font></strong></td>
    <td width="76" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">��Ʒ��˾</font></strong></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����ʱ��</font></strong></td>
    <td colspan="2" width="76" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="ȫ" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="��" style="HEIGHT: 18px;WIDTH: 16px;">	</td>
  </tr>
  <% ProductsList() %>
  </form>
</table>
</BODY>
</HTML>
<%
'-----------------------------------------------------------
function ProductsList()
  dim idCount'��¼����
  dim pages'ÿҳ����
      pages=20
  dim pagec'��ҳ��
  dim page'ҳ��
      page=clng(request("Page"))
  dim pagenc'ÿҳ��ʾ�ķ�ҳҳ������=pagenc*2+1
      pagenc=2
  dim pagenmax'ÿҳ��ʾ�ķ�ҳ�����ҳ��
  dim pagenmin'ÿҳ��ʾ�ķ�ҳ����Сҳ��
  dim datafrom'���ݱ���
      datafrom="NwebCn_Products"
  dim datawhere'��������
      if Result="Search" then
	     datawhere="where ( ProductName like '%" & Keyword &_
		           "%') and AddTime >= #" & StartDate & " # and AddTime <= #" & EndDate & "#"
	  else
	    if SortPath<>"" then'�Ƿ�鿴�ķ����Ʒ
		  datawhere="where charindex(SortPath,'"&SortPath&"')>0 "
        else
		  datawhere=""
		end if
	  end if
  dim sqlid'��ҳ��Ҫ�õ���id
  dim Myself,PATH_INFO,QUERY_STRING'��ҳ��ַ�Ͳ���
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'�������� asc,desc
      taxis="order by id desc"
  dim i'����ѭ��������
  dim rs,sql'sql���
  '��ȡ��¼����
  sql="select count(ID) as idCount from ["& datafrom &"]" & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
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
      if rs("ViewFlag") then
        Response.Write "<td nowrap><font color='blue'>��</font></td>" & vbCrLf
      else
        Response.Write "<td nowrap><font color='red'>��</font></td>" & vbCrLf
	  end if
      Response.Write "<td nowrap>"&rs("ProductNo")&"</td>" & vbCrLf
	  if StrLen((rs("ProductName")))>33 then
        Response.Write "<td nowrap title='���&#13;"&SortText(rs("SortID"))&"&nbsp;|&nbsp;"&rs("SortPath")&"'>"&StrLeft(rs("ProductName"),30)&"</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='���&#13;"&SortText(rs("SortID"))&"&nbsp;|&nbsp;"&rs("SortPath")&"'>"&rs("ProductName")&"</td>" & vbCrLf
      end if 
      Response.Write "<td nowrap><font color='blue'>" & vbCrLf
      if rs("NewFlag")=1 then Response.Write "<font color='red'>��</font>"
	  if rs("CommendFlag")=1 then Response.Write "<font color='green'>��</font>"
	  Response.Write "</td>"
      Response.Write "<td nowrap>"&rs("px")&"</td>" & vbCrLf
	  if StrLen((rs("Maker")))>13 then
        Response.Write "<td nowrap title='��Ʒ��˾��&#13;"&rs("Maker")&"'>"&StrLeft(rs("Maker"),10)&"</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='��Ʒ��˾��&#13;"&rs("Maker")&"'>"&rs("Maker")&"</td>" & vbCrLf
      end if 


      Response.Write "<td nowrap>"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td width='40' nowrap><a href='ProductEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""�޸Ĳ�Ʒ��Ϣ"")'><font color='#330099'>�޸�</font></a></td>" & vbCrLf
      Response.Write "<td width='22' nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='8' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='ɾ����ѡ' onClick='ConfirmDel(""�����Ҫɾ����Щ��Ʒ��"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='12' nowrap  bgcolor='#EBF2F9'>���޲�Ʒ��Ϣ</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='10' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
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
'���ɽڵ�����·��--------------------------
Function TextPath(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_ProductSort where ID="&ID
  rs.open sql,conn,1,1
  TextPath="&nbsp;->&nbsp;<a href=ProductList.asp?SortID="&rs("ID")&"&SortPath="&rs("SortPath")&">"&rs("SortName")&"</a>"
  if rs("ParentID")<>0 then TextPath rs("ParentID")
  response.write(TextPath)
End Function
%>
<%
'�����������--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_ProductSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>
