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
<TITLE>�����б�</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<link rel="stylesheet" href="Images/FilesCss.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" type="text/javascript" src="../Script/jquery-1.8.2.js"></script>
<style type="text/css">
<!--
.STYLE3 {color: #FFFFFF; font-weight: bold; }
-->
</style>

<script language="javascript">
function act1()  
{
var msg = "ȷ����ѡ����Ϊ�Ѿ�������";
if (confirm(msg)==true){
     document.formD.action="ChangStatep.asp?State=�Ѿ�����";  
     document.formD.submit();
}else{ 
return false; 
} 
}  
function act2()  
{
var msg = "�����Ҫɾ����Щ������Ϣ��";
if (confirm(msg)==true){
     document.formD.action="DelContent.asp?Result=Order";
     document.formD.submit();
}else{
return false; 
}
}
function act3()  
{
var msg = "ȷ����ѡ����ΪǮ���ѷ���";
if (confirm(msg)==true){
     document.formD.action="ChangStatep.asp?State=Ǯ���ѷ�";  
     document.formD.submit();  
}else{ 
return false; 
}
}  
</script>

</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="select_date.asp"-->
<%
if Instr(session("AdminPurview"),"|93,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if

'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<%
On Error Resume Next
dim Result,StartDate,EndDate,Keyword
Result=request.QueryString("Result")
StartDate=request.QueryString("StartDate")
EndDate=request.QueryString("EndDate")
Keyword=request.QueryString("Keyword")
 
function PlaceFlag()
  dim states
  states=Trim(Request("State"))
  if Result="Search" then
    Response.Write "�������б�&nbsp;->&nbsp;����&nbsp;->&nbsp;����ʱ��[<font color='red'>"&StartDate&"��"&EndDate&"</font>]���ؼ���[<font color='red'>"&Keyword&"</font>]"
  else
    Response.Write "�������б�&nbsp;->&nbspȫ��"
  end if
  response.Write("<select name='State' id='State' size='1' style='margin-left:10px;' onchange='Evern_Change();'>")
  	response.Write("<option value='NULL'>--ȫ��--</option>")
	
  	if states="δ����" then
		response.Write("<option value='δ����' selected>δ����</option>")
	else
		response.Write("<option value='δ����'>δ����</option>")
	end if
	
	if states="�����Ѹ�" then
		response.Write("<option value='�����Ѹ�' selected>�����Ѹ�</option>")
	else
		response.Write("<option value='�����Ѹ�'>�����Ѹ�</option>")
	end if
	
	if states="Ǯ���ѷ�" then
		response.Write("<option value='Ǯ���ѷ�' selected>Ǯ���ѷ�</option>")
	else
		response.Write("<option value='Ǯ���ѷ�'>Ǯ���ѷ�</option>")
	end if
	
	if states="���ܵ���" then
		response.Write("<option value='���ܵ���' selected>���ܵ���</option>")
	else
		response.Write("<option value='���ܵ���'>���ܵ���</option>")
	end if
	
	if states="�Ѿ�����" then
		response.Write("<option value='�Ѿ�����' selected>�Ѿ�����</option>")
	else	
		response.Write("<option value='�Ѿ�����'>�Ѿ�����</option>")
	end if
	
	if states="�ն�δ��" then
		 response.Write("<option value='�ն�δ��' selected>���и���</option>")
	else	
		response.Write("<option value='�ն�δ��'>���и���</option>")
	end if
	if states="δ����" then
		 response.Write("<option value='δ����' selected>δ����</option>")
	else	
		response.Write("<option value='δ����'>δ����</option>")
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
  response.Write("</select>")
end function  
%>
<script language="javascript">
<!--
function Evern_Change()
{
	var state=document.getElementById("State");
	window.location.href="OrderList.asp?State="+state.value;
}
-->
</script>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" ><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>������Ϣ���鿴���޸ģ��ظ�������Ϣ��ص�����</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center"   bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="Search.asp?Result=Order">
          <td > ������������
<%
	if Result="Search" then
		Response.Write "<input name=""start_date"" type=""text"" class=""textfield"" value="&StartDate&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date"" />��<input name=""end_date"" type=""text"" class=""textfield"" value="&EndDate&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date2"" />"
	else
		Response.Write "<input name=""start_date"" type=""text"" class=""textfield"" value="&dateadd("yyyy",-1,date())&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date"" />��<input name=""end_date"" type=""text"" class=""textfield"" value="&date()&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date2"" />"
	end if
%>
            <!--<script language=javascript> 
          var myDate=new dateSelector(); 
          myDate.year--; 
		  myDate.date; 
          myDate.inputName='start_date';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name�� 
          myDate.display(); 
          </script>
          &nbsp;��
          <script language=javascript> 
          myDate.year++; 
		  //myDate.date++;
          myDate.inputName='end_date';  //ע����������������name��ͬһҳ�е���������򣬲��ܳ����ظ���name�� 
          myDate.display(); 
          </script>-->
          &nbsp;&nbsp;�ؼ��֣�<input name="Keyword" type="text" class="textfield" value="<%=Keyword%>" size="18" />
          <input name="submitSearch" type="submit" class="button" value="����" />
          </td>
        </form>
        <td align="right" >�鿴��<a href="OrderList.asp" onClick='changeAdminFlag("������Ϣ�б�")'>���ж�����Ϣ</a></td>
      </tr>
    </table>
	</td>    
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"><%PlaceFlag()%><span style="margin-left:20px;">
    	<select name="KD_TYPE" id="KD_TYPE" >
        	<option value="0" >-��ѡ��-</option>
        	<option value="1" >��ͨ</option>
        	<option value="2" >EMS</option>
        	<option value="3" >EMS����</option>
        	<option value="4" >ZJS����</option>
        </select><input type="button" name='Excle' id="Excle" value="����Excle�ļ�" onClick="Create_ExcelFile();"> ��ѡ����״̬����ѡ���ںͿ�ݷ�ʽ</span>
	<form name="formxian" method="post" action="OrderList.asp" style="margin:0px;display:inline;">
	&nbsp;��ʾ&nbsp;<input name="num" type="text" class="textfield" size="6" maxlength="6" onkeyup="value=value.replace(/[^\d]/g,'')" />&nbsp;��
	<input name="submitxian" type="submit" class="button" value="ȷ��" />
	</form>
	</td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1" style="table-layout:automatic;word-wrap:break-word;">
  <form method="post" action="" name="formD" style="margin:0px;display:inline;">
  <tr>
    <td width="60" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>������</strong></font></td>
    <%if Instr(session("AdminPurview"),"|316,")<>0 then%>
    <td width="106" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>��Դ</strong></font></td>
    <%end if%>
    <td width="90" align="center" nowrap bgcolor="#8DB5E9"><span class="STYLE3">��������</span></td>
    <td width="70" align="center" bgcolor="#8DB5E9"><span class="STYLE3">��ϵ�绰</span></td>
    <td width="90" align="center"  bgcolor="#8DB5E9"><span class="STYLE3">�绰������</span></td>
   <!--
    <td width="96" align="center"  bgcolor="#8DB5E9"><span class="STYLE3">��ϵ�ֻ�</span></td>
   -->
   <td width="90" align="center" nowrap bgcolor="#8DB5E9"><span class="STYLE3">֧����ʽ</span></td>
    <td width="110" align="center"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����ʱ��</font></strong></td>
    <td width="220" align="center"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">��ַ ����</font></strong></td>
   <!--
    <td align="center"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">IP��ַ</font></strong></td>
   -->
    <td width="90" align="center"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">IP������</font></strong></td>
    <td width="116"  align="center" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">״ ̬</font></strong></td>
    <td colspan="2" width="90"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSelect" value="ȫ" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="��" style="HEIGHT: 18px;WIDTH: 16px;">	</td>
  </tr>
  <% OrderList() %>
  </form>
</table>
</BODY>
</HTML>
<%
'-----------------------------------------------------------
function OrderList()
On Error Resume Next

if request.Form("num") <> "" then
	session("num") = request.Form("num")
end if
dim num
num=session("num")'����POST����
  dim pages'ÿҳ����
if num="" then
      pages=100
else
      pages=num
end if

  dim idCount'��¼����

  dim pagec'��ҳ��
  dim page'ҳ��
      page=clng(request("Page"))
  dim pagenc'ÿҳ��ʾ�ķ�ҳҳ������=pagenc*2+1
      pagenc=2
  dim pagenmax'ÿҳ��ʾ�ķ�ҳ�����ҳ��
  dim pagenmin'ÿҳ��ʾ�ķ�ҳ����Сҳ��
  dim datafrom'���ݱ���
      datafrom="NwebCn_Order"
  dim datawhere'��������
  
      if Result="Search" then
	  
	     datawhere="where ( o.ProductName like '%" & Keyword &_
		           "%' or o.ProductNo like '%" & Keyword &_
		           "%' or o.Linkman like '%" & Keyword &_
		           "%' or o.Tel like '%" & Keyword &_
		           "%' or o.Company like '%" & Keyword &_
		           "%' or o.Address like '%" & Keyword &_
		           "%')  and (o.addtime between '" & StartDate & "' and '" & Cdate(EndDate)+1 & "')"
	  else
        datawhere=" where o.fax=0 and (o.blacklist=0 or o.blacklist=2) "
	  end if
	  
	  if Trim(Request("State"))<>"NULL" and Trim(Request("State"))<>"" then 
	  	if Trim(Request("State"))<>"������" then	  
			if datawhere="" then
				datawhere="where o.State='"&Trim(Request("State"))&"'"
			else
				if Trim(Request("State"))="δ����" then
				datawhere=datawhere&" and (o.State='"&Trim(Request("State"))&"' or o.State=NULL)"
				else
				datawhere=datawhere&" and o.State='"&Trim(Request("State"))&"'"
				end if
			end if
		else
			if datawhere="" then
				datawhere=" where o.HuoDao_FuKuan=1 and (o.State is Null)"
			else
				datawhere=datawhere&" and o.HuoDao_FuKuan=1 and (o.State is Null)"
			end if
		end if
	  end if
	'  datawhere=datawhere& " order by addtime desc "
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
      taxis="order by o.addtime desc"
  dim i'����ѭ��������
  dim rs,sql'sql���
  '��ȡ��¼����
  sql="select count(o.ID) as idCount from "& datafrom &" as o " & datawhere

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
    sql="select o.id from "& datafrom &" as o " & datawhere & taxis

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
    sql="select o.*,ae.ADS_Name as LinkName from "& datafrom &" as o Left Join NwebCn_Ads_effect as ae on o.ADS_Link  = ae.id where o.id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    while(not rs.eof)'������ݵ����

dim ipoder,IPgo
if rs("ipto") = "������ַ - IALVIN.CN" then
	IPgo = "������ַ"
else
	IPgo = rs("ipto")
end if
if rs("ipaddress") = "112.195.133.10" then
	IPgo = "�ͷ��µ�"
else
	ipoder = rs("ipaddress")
end if

	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
    
      Response.Write "<td >"&Guest(rs("MemID"),rs("Linkman"))&"</td>" & vbCrLf
      if Instr(session("AdminPurview"),"|316,")<>0 then
	  Response.Write "<td >"&rs("LinkName")&"</td>" & vbCrLf
	  end if
	  if StrLen(rs("Amount"))>50 then
        Response.Write "<td title="&rs("Amount")&" >"&Replace(replace(Print(rs("Amount")),"�����","(B)"),"����0��","")&"</td>" & vbCrLf
      else
        Response.Write "<td title="&rs("Amount")&" >"&Replace(Replace(replace(Print(rs("Amount")),"�����","(B)"),"һ��0�С�",""),"������0��","")&"</td>" & vbCrLf
      end if
	  
	  dim sms_states
	  if rs("sms_states")=true then
	    sms_states = "<font color='red' title='�ѷ���������'>��</font>"
		else
	    sms_states = "<font color='#333' title='δ����������'>��</font>"
	  end if

	'��ѯ�����ظ�
	'dim rstel
	'sql="SELECT Tel FROM NwebCn_Order WHERE Tel='"&rs("Tel")&"'"
	'set rstel=server.createobject("adodb.recordset")
	'rstel.open sql,conn,1,1

	'���ε绰����
	dim LastStr
	LastStr=Left(rs("Tel"),0)&"********"&right(rs("Tel"),3)

	if Instr(session("AdminPurview"),"|121,")>0 then 
'if rstel.recordcount >= 2 then
'		Response.Write "<td><font color='#0000ff'>"&rs("Tel")&"</font></td>" & vbCrLf
'	else
	if rs("blacklist") = 2 then
		Response.Write "<td>"&rs("Tel")&"&nbsp;<font color='#ff0000'>������</font></td>" & vbCrLf
	else
		Response.Write "<td>"&rs("Tel")&"</td>" & vbCrLf
	end if
'end if
	else
	if rs("blacklist") = 2 then
		Response.Write "<td>"&LastStr&"&nbsp;<font color='#ff0000'>������</font></td>" & vbCrLf
	else
		Response.Write "<td>"&LastStr&"</td>" & vbCrLf
	end if
	end if

	dim telsheng
	if Instr(rs("Telto"),",") >0 then
		telsheng = split(rs("Telto"),",")
	else
		telsheng = split(rs("Telto")," ")
	end if

	Response.Write "<td width='90' style='word-break:break-all;word-wrap:break-word;'>"
	if Instr(IPgo,telsheng(1)) > 0 or Instr(rs("Address"),telsheng(1)) > 0 then
	Response.Write HighLight(rs("Telto"),telsheng(1))
	else
	Response.Write rs("Telto")
	end if
	Response.Write "</td>"

	 'Response.Write "<td width='80'>"&rs("Telto")&"</td>" & vbCrLf

	  Response.Write("<td align='center'>"&vbcrlf)
	  	Dim ZiFu_FS
		ZiFu_FS=Split(rs("Remark"),"|")
		Response.Write(ZiFu_FS(1))
		Response.Write(ZiFu_FS(2))
		Response.Write("Ԫ")
	  Response.Write("</td>")
      'Response.Write "<td title='"&Rs("ipaddress")&"|"&LookAdd(Rs("ipaddress"))&"' >"&rs("AddTime")&"</td>" & vbCrLf
	  Response.Write "<td  >"&rs("AddTime")&"</td>" & vbCrLf

	Response.Write "<td>"
	if HighLight(IPgo,telsheng(1)) = "" then
	Response.Write rs("Address")&"("&rs("zipcode")&")"
	else
	Response.Write HighLight(rs("Address"),telsheng(1))&"("&rs("zipcode")&")"
	end if
	Response.Write "</td>"
	'Response.Write "<td>"&rs("Address")&"("&rs("zipcode")&")</td>" & vbCrLf

	Response.Write "<td width='90' style='word-break:break-all;word-wrap:break-word;' title=""IP��ַ��"&rs("ipaddress")&""">"
	Response.Write HighLight(IPgo,telsheng(1))
	Response.Write "</td>"
	'Response.Write "<td width='80' title=""IP��ַ��"&rs("ipaddress")&""">"&IPgo&"</td>" & vbCrLf '��ʾIP��ַ

	  Response.Write("<td align='center' style='color:#ff0000;'>"&vbcrlf)
	  	'//�����޸ĵĴ���
		if Instr(session("AdminPurview"),"|314,")>0 then
		if rs("State")<>"" and rs("State") <>"δ����" then
			if instr(rs("State"),"�����Ѹ�") > 0 then
				response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=Ǯ���ѷ�'>Ǯ���ѷ�</a> | ")
				response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=�Ѿ�����'>�Ѿ�����</a> | ")
				response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=δ����'>����״̬</a>")
			else
				response.Write(Replace(rs("State"),"�ն�δ��","���и���"))
				response.Write("|<a href='ChangState.asp?ID="&rs("ID")&"&State=δ����'>����״̬</a>")
			end if
		else
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=δ����'>δ����</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=���ܵ���'>���ܵ���</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=�����Ѹ�'>�����Ѹ�</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=Ǯ���ѷ�'>Ǯ���ѷ�</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=�Ѿ�����'>�Ѿ�����</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=�ն�δ��'>���и���</a>")
			 
		end if
		else
		if rs("State")="" then
	    response.Write("δ����")
		else
	    response.Write(rs("State"))
		end if
	   end if
		'//�ڶ����޸ĺ�Ĵ���
		'if rs("State")="δ����" then
			'response.Write("<a href='fukuan.asp?id="&rs("id")&"'><font color='#ff0000'>�Ѹ���</font></a>")
		'elseif rs("State")="�Ѹ���" then
			'response.Write("<a href='fahuo.asp?id="&rs("id")&"'><font color='#ff0000'>����</font></a>")
		'elseif rs("State")="���ѷ�" then
			'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
		'elseif rs("State")="δ����" then
			'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
		'elseif rs("State")="��δ�յ�" then
			'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
		'else
			'if rs("HuoDao_FuKuan") then
				'if rs("State")="" or isnull(rs("State")) then
					'response.Write("<a href='HuoDaoFk.asp?id="&rs("ID")&"&Action=true'>��������</font>")
					'response.Write("<br />")
					'response.Write("<a href='HuoDaoFk.asp?id="&rs("ID")&"&Action=false'>���ܷ���</a>")
				'else
					'if rs("FuKuan") then
						'response.Write("<a href='fahuo.asp?id="&rs("ID")&"'>����</font>")
					'else
						'response.Write("���ܷ���<br>")
						'response.Write("<a href='HuoDaoFk.asp?id="&rs("ID")&"&Action=true'>��Ϊ��������</font>")
						
					'end if
				'end if
			'end if
		'end if
	  response.Write("</td>"&vbcrlf) 	  
      Response.Write "<td colspan=2 align='center'><a href='OrderEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""�鿴�ظ�������Ϣ"")'><font color='#330099'>�鿴</font></a>&nbsp;"

	  dim X
	  if Instr(session("AdminPurview"),"|306,")=0 then
	   X="disabled"
	  else
	   X=""
	  end if 
	  response.Write("&nbsp;<input "&x&" name='selectID' type='checkbox' value='"&rs("ID")&"' onclick='' style='HEIGHT: 13px;WIDTH: 13px;'>")
	  dim KDFS
	  KDFS = 0
	  IF rs("KDFS")="" then
	   KDFS = 0
	  else
	   KDFS = Cint(rs("KDFS"))
	  end if
	  %>
	  	<!--<select name="KDFS" onChange="ChangeKDFS(<%=rs("ID")%>,this.value)" >
        	<option value="0" <%if KDFS=0 then%>selected<%end if%>>-��ѡ��-</option>
        	<option value="1" <%if KDFS=1 then%>selected<%end if%>>��ͨ</option>
        	<option value="2" <%if KDFS=2 then%>selected<%end if%>>EMS</option>
        	<option value="3" <%if KDFS=3 then%>selected<%end if%>>EMS����</option>
        	<option value="4" <%if KDFS=4 then%>selected<%end if%>>ZJS����</option>
        </select>-->
	  <%

	  if Instr(session("AdminPurview"),"|313,")<>0 then 
      response.Write("<a href='OrderDel.asp?Result=Modify&ID="&rs("ID")&"'>ɾ��</a>")
if rs("blacklist") = 2 then
      response.Write("&nbsp;<a href='huifu.asp?id="&Rs("ID")&"&zt=bl'>�ָ�</a>")
else
      response.Write("&nbsp;<a href='BlackListDel.asp?Result=Modify&ID="&rs("ID")&"'>����</a>")
end if
	  end if

	  Response.Write("</td>")
	  
	  Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='9' bgcolor='#EBF2F9'></td>" & vbCrLf

 Response.Write "<td colspan='2' style='padding:10px 0 10px 10px;' bgcolor='#EBF2F9'><input type='submit' class='button' value='�Ѿ�����' onClick='act1();'><br /><br /><input type='submit' class='button' value='Ǯ���ѷ�' onClick='act3();'>" & vbCrLf
if session("AdminId") = 1 then
    Response.Write "<br /><br /><input "&x&" name='submitDelSelect' type='button' class='button' id='submitDelSelect' value='ɾ����ѡ' onClick='act2();'>" & vbCrLf
end if
    Response.Write "</td></tr>" & vbCrLf
   
  else
    response.write "<tr><td height='50' align='center' colspan='11'   bgcolor='#EBF2F9'>���޶�����Ϣ</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='11'   bgcolor='#D7E4F7'>" & vbCrLf
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
	  response.write ("[<a href="& myself &"Page="& i &"&State="&Trim(Request("State"))&">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write ("&nbsp;<a href='"& myself &"Page="& page+(pagenc*2+1) &"&State="&Trim(Request("State"))&"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>8</font></a>&nbsp;") '���ҳ�����ֵС����ҳ������ʾ(����)
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"&State="&Trim(Request("State"))&"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>:</font></a>&nbsp;") '���ҳ��С����ҳ������ʾ(���ҳ)	
  '���÷�ҳҳ�����===============================
  Response.Write "��������&nbsp;<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('ֻ������תĿ��ҳ��������������');this.value='"&Page&"';}"" style='HEIGHT: 18px;WIDTH: 40px;'  type='text' class='textfield' value='"&Page&"'>&nbsp;ҳ" & vbCrLf
  Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick=""GoPage2('"&Myself&"','"&Trim(Request.QueryString("State"))&"')"" value='GO'>" & vbCrLf
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

function Guest(ID,Linkman)
On Error Resume Next
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    Guest=Linkman
  else
    Guest="<font color='green'>��Ա��</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""ǰ̨��Ա����"")'>"&Linkman&"</a>"
  end if
  rs.close
  set rs=nothing
end function 
function PringText(Remark)
On Error Resume Next
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

function Print(Amount)
On Error Resume Next
	dim str,i,str1,aa
	str1=""
	aa=replace(Amount,"��","(")
	aa=replace(aa,"��",")")
	str=split(aa,"|")
	'if ubound(str)>0 then
	for i=0 to ubound(str)
		if i>0 then str1=str1&"��"
		if str1="" then
			str1=Mid(str(i),1,instr(str(i),"(")-1)
		else
			str1=str1&Mid(str(i),1,instr(str(i),"(")-1)
		end if
		str1=str1&Mid(str(i),instr(str(i),"(")+1,(instr(str(i),")"))-(instr(str(i),"(")+1))&"��"
	next
	Print=str1
	'else
	'Print=Amount
	'end if
end function
 Function LookAdd(Sip)
  Dim Str1,Str2,Str3,Str4
  Dim Num
  Dim Irs,sql
  If IsNumeric(Left(sip,2)) Then
    If Sip="127.0.0.1" Then sip="192.168.0.1"
      Str1=Left(Sip,InStr(Sip,".")-1)
      Sip=Mid(Sip,InStr(Sip,".")+1)
      Str2=Left(Sip,InStr(Sip,".")-1)
      Sip=Mid(Sip,InStr(Sip,".")+1)
      Str3=Left(Sip,InStr(Sip,".")-1)
      Str4=Mid(Sip,InStr(Sip,".")+1)
  If IsNumeric(Str1)=0 Or isNumeric(Str2)=0 Or isNumeric(Str3)=0 Or isNumeric(Str4)=0 Then
   Else
    num=CInt(Str1)*256*256*256+CInt(Str2)*256*256+CInt(Str3)*256+CInt(Str4)-1
    Dim adb,aConnStr,AConn
    adb = "../DATAbase/ip.mdb"
    aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(adb)
    Set AConn = Server.CreateObject("ADODB.Connection")
    aConn.Open aConnStr
    sql="select country from IPTABLE where StartIPnum <="&num&" and EndIPnum >="&num
    Set irs=AConn.Execute(sql)
    If irs.eof And irs.bof Then 
     LookAdd="�й�"
    Else
     Do While Not irs.eof
      LookAdd=LookAdd & Irs(0) 
     Irs.MoveNext
     Loop
    End If
    Irs.Close
    Set Irs=nothing
    Set AConn=Nothing
   End If
  End If
 End Function 

function HighLight(yu,bi)
Dim regEx
Set regEx = New RegExp
regEx.IgnoreCase = True
regEx.Global = True
regEx.Pattern = "("&bi&")"
HighLight = regEx.Replace(yu,"<span style='color:#ff6600'>$1</span>")
End function

%>
<script language="javascript">
<!--
function GoPage2(Myself,str2)
{
	window.location.href=Myself+"Page="+document.formD.SkipPage.value+"&State="+str2;
}
function Create_ExcelFile()
{	
	var iframe=document.createElement("iframe");
	iframe.style.width=0;
	iframe.style.height=0;
	iframe.style.border=0;
	
	var UrlValue=window.location.href;
	var GetValue;
	UrlValue=UrlValue.slice(UrlValue.lastIndexOf("/")+1,UrlValue.length);
	if(UrlValue.indexOf("?")!=-1)
	{
		GetValue=UrlValue.slice(UrlValue.indexOf("?")+1,UrlValue.length)
	} 
	var start_date = $("#DS_start_date").val();
	var end_date = $("#DS_end_date").val();
	var KDFS = $("#KD_TYPE").val();
	var sta = $("#State").val();
	var url="expXLS.asp?s="+start_date+"&e="+end_date+"&f="+KDFS+"&sta="+sta;
	iframe.src=url;
	window.document.body.appendChild(iframe);
	
}
function ChangeKDFS(order_id,value)
{
	url = "changeKDFS.asp?order_id="+order_id+"&f="+value;
	$.get(url,null,responseFun,null);
}
function responseFun(result)
{
	//alert(result)
}
-->

</script>
<div style="display:none;"><script src="http://s21.cnzz.com/stat.php?id=4968983&web_id=4968983" language="JavaScript"></script></div>