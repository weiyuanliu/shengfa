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
<TITLE>�༭����</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<%
call CreateEditor("Content")
%>

</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|23,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,NewsName,ViewFlag,SortName,SortID,SortPath
dim GroupID,GroupIdName,Exclusive,NoticeFlag,Source,Content,px,CommendFlag,smallpic,bigpic,daodu
ID=request.QueryString("ID")
call NewsEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>���ż���������鿴����ӣ��޸ģ�ɾ��������Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="NewsEdit.asp?Result=Add" onClick='changeAdminFlag("��Ӳ�Ʒ��Ϣ")'>���������Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="NewsList.asp" onClick='changeAdminFlag("��Ʒ�б�")'>�鿴����������Ϣ</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="NewsEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">�������ƣ�</td>
        <td><input name="NewsName" type="text" class="textfield" id="NewsName" style="WIDTH: 240;" value="<%=NewsName%>" maxlength="100">&nbsp;��ʾ��<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1" <%if ViewFlag or Result="Add" then response.write ("checked")%>>&nbsp;*&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">�������</td>
        <td><input name="SortName" type="text" class="textfield" id="SortName" value="<%=SortName%>" style="WIDTH: 240;background-color:#EBF2F9;" readonly>&nbsp;<a href="javaScript:OpenScript('SelectSort.asp?Result=News',500,500,'')"><img src="Images/Select.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">������֣�</td>
        <td><input name="SortID" type="text" class="textfield" id="SortID" style="WIDTH: 40;background-color:#EBF2F9;" value="<%=SortID%>" readonly><input name="SortPath" type="text" class="textfield" id="SortPath" style="WIDTH: 200;background-color:#EBF2F9;" value="<%=SortPath%>" readonly>&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">������Դ��</td>
        <td><input name="Source" type="text" class="textfield" style="WIDTH: 240;" value="<%=Source%>" maxlength="100"></td>
      </tr>
	    <tr>
        <td height="20" align="right">����</td>
        <td><input name="px" type="text" class="textfield" style="WIDTH: 60px;" value="<%=px%>" maxlength="100">*ֻ����������</td>
      </tr>

      <tr>
        <td height="20" align="right">�ꡡ���ǣ�</td>
        <td><input name="NoticeFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if NoticeFlag then response.write ("checked")%>>&nbsp;����&nbsp;&nbsp;<input name="CommendFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if CommendFlag then response.write ("checked")%>>&nbsp;�Ƽ�</td>
      </tr>
	        <tr>
        <td height="20" align="right">������ͼ��</td>
        <td><input name="BigPic" type="text" class="textfield" style="WIDTH: 240;" value="<%=BigPic%>" maxlength="100">
        &nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=BigPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"> �Ƽ�341*199 </a><a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)">�õ�ͼƬ������jpg��ʽ</a></td>
      </tr>
      <tr>
        <td height="20" align="right">������ �� ͼ��</td>
        <td><input name="SmallPic" type="text" class="textfield" style="WIDTH: 240;" value="<%=SmallPic%>" maxlength="100">
        &nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"> �Ƽ�130*85 </a></td>
      </tr>
      <tr>
        <td height="20" align="right">�鿴Ȩ�ޣ�</td>
        <td><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> ����<input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��</td>
      </tr>
	  	    <tr>
        <td height="20" align="right">����������</td>
        <td> <input  name="daodu" size='120' class="textfield"  value="<%=daodu%>" />        ������200�ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">��Ϣ���ݣ�<br>
        <td style="padding:6px"><textarea name="Content" rows="30" class="textfield" id="Content" style="WIDTH: 86%;" ><%=Content%></textarea></td>
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
sub NewsEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭��Ʒ��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("NewsName")))<1 then
      response.write ("<script language=javascript> alert('��������Ϊ������Ŀ��');history.back();</script>")
      response.end
    end if
    if Result="Add" then '������Ʒ��Ϣ
	  sql="select * from NwebCn_News"
      rs.open sql,conn,1,3
      rs.addnew
      rs("NewsName")=trim(Request.Form("NewsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")="" and Request.Form("SortPath")="" then
        response.write ("<script language=javascript> alert('��ѡ���������࣡');history.back();</script>")
        response.end
	  else
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  end if
	  rs("Source")=trim(Request.Form("Source"))
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  if Request.Form("NoticeFlag")=1 then
        rs("NoticeFlag")=Request.Form("NoticeFlag")
	  else
        rs("NoticeFlag")=0
	  end if
	  if Request.Form("CommendFlag")=1 then
        rs("CommendFlag")=Request.Form("CommendFlag")
	  else
        rs("CommendFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("Content")=trim(Request.Form("Content"))
	  rs("BigPic")=trim(Request.Form("BigPic"))	  
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  rs("daodu")=trim(Request.Form("daodu"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸Ĳ�Ʒ��Ϣ
      sql="select * from NwebCn_News where ID="&ID
      rs.open sql,conn,1,3
      rs("NewsName")=trim(Request.Form("NewsName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('��ѡ���������࣡');history.back();</script>")
        response.end
	  end if
	  rs("Source")=trim(Request.Form("Source"))
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  if Request.Form("NoticeFlag")=1 then
        rs("NoticeFlag")=Request.Form("NoticeFlag")
	  else
        rs("NoticeFlag")=0
	  end if
	   if Request.Form("commendFlag")=1 then
        rs("commendFlag")=Request.Form("commendFlag")
	  else
        rs("commendFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("Content")=trim(Request.Form("Content")) 
	  rs("BigPic")=trim(Request.Form("BigPic"))	  
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  
	  rs("daodu")=trim(Request.Form("daodu"))
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭������Ϣ��');changeAdminFlag('�����б�');location.replace('NewsList.asp');</script>"
  else '��ȡ��Ʒ��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_News where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("���ݿ��ȡ��¼����")
        response.end
      end if
	  NewsName=rs("NewsName")
	  ViewFlag=rs("ViewFlag")
	  SortName=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  SortPath=rs("SortPath")
	  Source=rs("Source")
	  px=rs("px")
	  NoticeFlag=rs("NoticeFlag")
	  CommendFlag=rs("CommendFlag")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
      Content=rs("Content")
	  BigPic=rs("BigPic")
	   daodu=rs("daodu")
	  SmallPic=rs("SmallPic")
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
<%
'�����������--------------------------
Function SortText(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_NewsSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>