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
<TITLE>�༭��Ʒ</TITLE>
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
if Instr(session("AdminPurview"),"|33,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,ProductName,ViewFlag,SortName,SortID,SortPath
dim ProductNo,Price,Px,Maker,CommendFlag,NewFlag,GroupID,GroupIdName,Exclusive,PriceText
dim BigPic,SmallPic,Content,Price2
ID=request.QueryString("ID")
call ProductEdit() 
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��Ʒ����������鿴����ӣ��޸ģ�ɾ����Ʒ��Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="ProductEdit.asp?Result=Add" onClick='changeAdminFlag("��Ӳ�Ʒ��Ϣ")'>��Ӳ�Ʒ��Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="ProductList.asp" onClick='changeAdminFlag("��Ʒ�б�")'>�鿴���в�Ʒ��Ϣ</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="ProductEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">��Ʒ���ƣ�</td>
        <td><input name="ProductName" type="text" class="textfield" id="ProductName" style="WIDTH: 240;" value="<%=ProductName%>" maxlength="100">&nbsp;������<input name="ViewFlag" type="checkbox" style='HEIGHT: 13px;WIDTH: 13px;' value="1"  <%if ViewFlag  or Result="Add"  then response.write ("checked")%>>
&nbsp;*&nbsp;������3���ַ�</td>
      </tr>
      <tr>
        <td height="20" align="right">�������</td>
        <td><input name="SortName" type="text" class="textfield" id="SortName" value="<%=server.HTMLEncode(SortName)%>" style="WIDTH: 240;background-color:#EBF2F9;" readonly>&nbsp;<a href="javaScript:OpenScript('SelectSort.asp?Result=Products',500,500,'')"><img src="Images/Select.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">������֣�</td>
        <td><input name="SortID" type="text" class="textfield" id="SortID" style="WIDTH: 40;background-color:#EBF2F9;" value="<%=SortID%>" readonly><input name="SortPath" type="text" class="textfield" id="SortPath" style="WIDTH: 200;background-color:#EBF2F9;" value="<%=SortPath%>" readonly>&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">�ࡡ���ţ�</td>
        <td><input name="ProductNo" type="text" class="textfield" id="ProductNo" style="WIDTH: 240;" value="<%=ProductNo%>" maxlength="100">&nbsp;*&nbsp;�������ȷ�����޸�</td>
      </tr>
      <tr>
        <td height="20" align="right">��������۸�</td>
        <td><input name="Price" type="text" class="textfield" id="Price" style="WIDTH: 240;" value="<%=Price%>" maxlength="100"></td>
      </tr>
      <tr>
        <td height="20" align="right">�������۸�</td>
        <td><input name="Price2" type="text" class="textfield" id="Price2" style="WIDTH: 240;" value="<%=Price2%>" maxlength="100"></td>
      </tr>
      <tr>
        <td height="20" align="right">�����۸�˵����</td>
        <td><input name="PriceText" type="text" class="textfield" id="PriceText" style="WIDTH: 240;" value="<%=PriceText%>" maxlength="100"></td>
      </tr>
      <tr>
        <td height="20" align="right">�š�����</td>
        <td><input name="Px" type="text" class="textfield" id="Px" style="WIDTH: 60px;" value="<%=Px%>" maxlength="100"> ֻ����д����</td>
      </tr>	  
	  <tr>
        <td height="20" align="right">��Ʒ��˾��</td>
        <td><input name="Maker" type="text" class="textfield" id="Maker" style="WIDTH: 240;" value="<%=Maker%>" maxlength="100"></td>
      </tr>
      <tr>
        <td height="20" align="right">״����̬��</td>
        <td><input name="CommendFlag" type="checkbox" style="HEIGHT: 13px;WIDTH: 13px;" value="1" <%if CommendFlag then response.write ("checked")%>>
        ��ҳ�Ƽ���ʾ&nbsp;
        <input name="NewFlag" type="checkbox" value="1" style="HEIGHT: 13px;WIDTH: 13px;" <%if NewFlag then response.write ("checked")%>>
        (���ã��ڱ�ϵͳδռ��)</td>
      </tr>
      <tr>
        <td height="20" align="right">�鿴Ȩ�ޣ�</td>
        <td><select name="GroupID" class="textfield">
          <% call SelectGroup() %>
          </select>
          <input name="Exclusive" type="radio" value="&gt;="  <%if Exclusive="" or Exclusive=">=" then response.write ("checked")%>> ����<input type="radio"  <%if Exclusive="=" then response.write ("checked")%> name="Exclusive" value="=">ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��</td>
      </tr>
      <tr>
        <td height="20" align="right">��Ʒ��ͼ��</td>
        <td><input name="BigPic" type="text" class="textfield" style="WIDTH: 240;" value="<%=BigPic%>" maxlength="100">&nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=BigPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"></a></td>
      </tr>
      <tr>
        <td height="20" align="right">�� �� ͼ��</td>
        <td><input name="SmallPic" type="text" class="textfield" style="WIDTH: 240;" value="<%=SmallPic%>" maxlength="100">
        &nbsp;<a href="javaScript:OpenScript('UpFileForm.asp?Result=SmallPic',460,180)"><img src="Images/Upload.gif" width="30" height="16" border="0" align="absmiddle"> �Ƽ�130*88</a></td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">��ϸ���ܣ�<br>
        <td  style="padding:6px"><textarea name="Content" rows="30" class="textfield" id="Content" style="WIDTH: 86%;" ><%=Content%></textarea></td>
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
sub ProductEdit()
  dim Action,rsRepeat,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '����༭��Ʒ��Ϣ
    set rs = server.createobject("adodb.recordset")
    if len(trim(request.Form("ProductName")))<3 then
      response.write ("<script language=javascript> alert('��Ʒ����Ϊ������Ŀ��');history.back(-1);</script>")
      response.end
    end if
    if Result="Add" then '������Ʒ��Ϣ
	  sql="select * from NwebCn_Products"
      rs.open sql,conn,1,3
      rs.addnew
      rs("ProductName")=trim(Request.Form("ProductName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")="" and Request.Form("SortPath")="" then
        response.write ("<script language=javascript> alert('��ѡ���������࣡');history.back(-1);</script>")
        response.end
	  else
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  end if
      set rsRepeat = conn.execute("select ProductNo from NwebCn_Products where ProductNo='" & trim(Request.Form("ProductNo")) & "'")
      if not (rsRepeat.bof and rsRepeat.eof) then '�жϴ˲�Ʒ����Ƿ����
        response.write "<script language=javascript> alert('" & trim(Request.Form("ProductNo")) & "�˲�Ʒ����Ѿ����ڣ��뻻һ����������ԣ�');history.back(-1);</script>"
        response.end
      else
	    rs("ProductNo")=trim(Request.Form("ProductNo"))
	  end if
	  rs("Price")=trim(Request.Form("Price"))
	   rs("Price2")=trim(Request.Form("Price2"))
	  rs("PriceText")=trim(Request.Form("PriceText"))
	  if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	  rs("Maker")=trim(Request.Form("Maker"))
	  if Request.Form("CommendFlag")=1 then
        rs("CommendFlag")=Request.Form("CommendFlag")
	  else
        rs("CommendFlag")=0
	  end if
	  if Request.Form("NewFlag")=1 then
        rs("NewFlag")=Request.Form("NewFlag")
	  else
        rs("NewFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("BigPic")=trim(Request.Form("BigPic"))	  
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  rs("Content")=Request.Form("Content")
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '�޸Ĳ�Ʒ��Ϣ
      sql="select * from NwebCn_Products where ID="&ID
      rs.open sql,conn,1,3
      rs("ProductName")=trim(Request.Form("ProductName"))
	  if Request.Form("ViewFlag")=1 then
        rs("ViewFlag")=Request.Form("ViewFlag")
	  else
        rs("ViewFlag")=0
	  end if
	  if Request.Form("SortID")<>"" and Request.Form("SortPath")<>"" then
	    rs("SortID")=Request.Form("SortID")
		rs("SortPath")=Request.Form("SortPath")
	  else
        response.write ("<script language=javascript> alert('��ѡ���������࣡');history.back(-1);</script>")
        response.end
	  end if
	  rs("ProductNo")=trim(Request.Form("ProductNo"))
	  rs("PriceText")=trim(Request.Form("PriceText"))
	  rs("Price")=trim(Request.Form("Price"))
	  rs("Price2")=trim(Request.Form("Price2"))
	  if  not isnumeric(trim(Request.Form("Px"))) then
	   rs("Px")=0
	   else
	   if isnumeric(trim(Request.Form("px"))) then
	  rs("px")=trim(Request.Form("px"))
	  else
	  rs("px")=0
	  end if
	   end if
	  rs("Maker")=trim(Request.Form("Maker"))
	  if Request.Form("CommendFlag")=1 then
        rs("CommendFlag")=Request.Form("CommendFlag")
	  else
        rs("CommendFlag")=0
	  end if
	  if Request.Form("NewFlag")=1 then
        rs("NewFlag")=Request.Form("NewFlag")
	  else
        rs("NewFlag")=0
	  end if
      GroupIdName=split(Request.Form("GroupID"),"���橾")
	  rs("GroupID")=GroupIdName(0)
	  rs("Exclusive")=trim(Request.Form("Exclusive"))
	  rs("BigPic")=trim(Request.Form("BigPic"))	  
	  rs("SmallPic")=trim(Request.Form("SmallPic"))
	  rs("Content")=Request.Form("Content")
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('�ɹ��༭��Ʒ��Ϣ��');changeAdminFlag('��Ʒ�б�');location.replace('ProductList.asp');</script>"
  else '��ȡ��Ʒ��Ϣ
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Products where ID="& ID
      rs.open sql,conn,1,1
      if rs.bof and rs.eof then
        response.write ("���ݿ��ȡ��¼����")
        response.end
      end if
	  ProductName=rs("ProductName")
	  ViewFlag=rs("ViewFlag")
	  SortName=SortText(rs("SortID"))
	  SortID=rs("SortID")
	  PriceText=rs("PriceText")
	  SortPath=rs("SortPath")
	  ProductNo=rs("ProductNo")
      Price=rs("Price")
	  Price2=rs("Price2")
	  Px=rs("Px")
	  Maker=rs("Maker")
	  CommendFlag=rs("CommendFlag")
	  NewFlag=rs("NewFlag")
	  GroupID=rs("GroupID")
	  Exclusive=rs("Exclusive")
	  BigPic=rs("BigPic")
	  SmallPic=rs("SmallPic")
      Content=rs("Content")
	  rs.close
      set rs=nothing 
	else
	  ProductNo="Pro"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
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
  sql="Select * From NwebCn_ProductSort where ID="&ID
  rs.open sql,conn,1,1
  SortText=rs("SortName")
  rs.close
  set rs=nothing
End Function
%>