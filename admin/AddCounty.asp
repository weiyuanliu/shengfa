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
<TITLE>����Ա�б�</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
Dim Action,ParetID
Action=Trim(Request.QueryString("Action"))
if Action="AddProvince" then Call SaveProvince()
%>

<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>��Ӷ����м���Ϣ</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AddCity.asp?Result=Add" onClick='changeAdminFlag("����м���Ϣ")'>����м���Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="County.asp" onClick='changeAdminFlag("�鿴�����м���Ϣ")'>�鿴�����м���Ϣ</a></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9">
    <table width="78%" border="0" cellpadding="5" cellspacing="0">
       <form name="AddProvince" id="AddProvince" method="post" action="AddCounty.asp?Action=AddProvince" onSubmit="return CheckValue();">
      <tr>
        <td height="25" align="right">����ʡ����</td>
        <td height="25">
        	<select name="ParentID" id="ParentID" size="1" onChange="Event_Change();">
            	<%Call GetParent()%>
            </select>        	��<span class="STYLE1">*����</span> </td>
      </tr>
      <tr>
        <td height="25" align="right">���������ƣ�</td>
        <td height="25">
        <select name="ParentID2" id="ParentID2" size="1">
            	<%Call GetParent2()%>
        </select>
        <span class="STYLE1">��*����</span>
        </td>
      </tr>
      <tr>
        <td width="21%" height="25" align="right">�أ����������ƣ�</td>
        <td width="79%" height="25"><input type="text" name="Content" id="Content">��
          <span class="STYLE1">*����</span></td>
      </tr>
      <tr>
        <td height="25" align="right">���ʱ�䣺</td>
        <td height="25"><input name="AddTime" type="text" id="AddTime" value="<%=now()%>"></td>
      </tr>
      <tr>
        <td height="25" align="right">����˳��</td>
        <td height="25"><input name="Px" type="text" id="Px" value="0"> ��
          *����д���֣�ԽС����Խǰ</td>
      </tr>
      
      <tr>
        <td height="40" align="right">&nbsp;</td>
        <td height="40"><label>
          <input type="submit" name="button" id="button" value="�� ��" style="margin-right:10px;">
          <input type="reset" name="button2" id="button2" value="�� ��">
        </label></td>
      </tr>
        </form>
    </table>
  
    </td>    
  </tr>
</table>
</body>
</html>
<%
sub SaveProvince()
	dim Content,AddTime,Px,ID,ParentID,ParentID2
	ID=Trim(Request.Form("ID"))
	Content=Trim(Request.Form("Content"))
	AddTime=Trim(Request.Form("AddTime"))
	Px=Trim(Request.Form("Px"))
	ParentID=Trim(Request.Form("ParentID"))
	ParentID2=Trim(Request.Form("ParentID2"))
	
	if Content="" or isnull(Content) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('����дʡ����Ϣ��');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	end if
	
	if Px="" or not(IsNumeric(Px)) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('����д��Ч��������Ϣ��');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
	end if
	
	if ParentID="" or isnull(ParentID) or ParentID="Null" then
		response.Write("<script language=javascript>"&vbcrlf) 
			response.Write("alert('û��ʡ����Ϣ��������Ӷ�����Ϣ���������ʡ����Ϣ��');"&vbcrlf)
			response.Write("window.location.href='AddProvince.asp?Result=Add';")
	 	response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	else
		if Not(IsNumeric(ParentID)) or isnull(ParentID) or ParentID="" then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('���ݳ������ܼ�����');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		end if
	end if
	
	if ParentID2="" or isnull(ParentID2) or ParentID2="Null" then
		response.Write("<script language=javascript>"&vbcrlf) 
			response.Write("alert('û��ʡ����Ϣ��������Ӷ�����Ϣ���������ʡ����Ϣ��');"&vbcrlf)
			response.Write("window.location.href='AddCity.asp?Result=Add';")
	 	response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	else
		if Not(IsNumeric(ParentID2)) or isnull(ParentID2) or ParentID2="" then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('���ݳ������ܼ�����');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		end if
	end if
	
	Dim rs,sql
	set rs=server.CreateObject("Adodb.Recordset")
	
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		sql="select * from County where Content='"&Content&"'"
	else
		sql="select * from County where Content='"&Content&"' and ParentID="&ParentID&" and ParentID2="&ParentID2&" and id not in("&id&")"
	end if
	rs.open sql,conn,1,1
	
	if not rs.eof and not rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script langauge=javascript>"&vbcrlf)
			response.Write("alert('�Բ��𣬲����ظ���ӣ�');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	end if
	
	rs.close()
	
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		Sql="Select top 1 * from County"
	else
		Sql="Select top 1 * from County where id="&ID
	end if
	rs.open sql,conn,1,3
	if ID="" or isnull(ID) or not(IsNumeric(ID)) then
		rs.addnew()
		rs("Content")=Content
		if AddTime<>"" then
			rs("AddTime")=AddTime
		else
			rs("AddTime")=Now()
		end if
		rs("ParentID")=ParentID
		rs("ParentID2")=ParentID2
		rs("Px")=Px
		rs.update()
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('�����ɹ���');"&vbcrlf)
			response.Write("window.location.href=document.referrer;")
		response.Write("</script>"&vbcrlf)
		response.End()
		exit sub
	else
		if rs.eof and rs.bof then
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('��¼δ�ҵ�������ʧ�ܣ�');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		else
			rs("Content")=Content
			if AddTime<>"" then
				rs("AddTime")=AddTime
			else
				rs("AddTime")=Now()
			end if
			rs("ParentID")=ParentID
			rs("ParentID2")=ParentID2
			rs("Px")=Px
			rs.update()
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert(���¼�¼�ɹ���);"&vbcrlf)
				response.Write("window.location.href=document.referrer;")
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		end if
	end if
end sub

sub GetParent()
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from Province order by px asc,id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<option value='Null'>������Ϣ</option>")	
	else
		ParetID=rs("ID")
		while not rs.eof 
			response.Write("<option value='"&rs("id")&"'>"&rs("Content")&"</option>")
			rs.movenext
		wend
	end if
	rs.close()
	set rs=Nothing
End sub

sub GetParent2()
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from City where ParentID="&ParetID&" order by px asc,id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<option value='Null'>������Ϣ</option>")	
	else
		while not rs.eof 
			response.Write("<option value='"&rs("id")&"'>"&rs("Content")&"</option>")
			rs.movenext
		wend
	end if
	rs.close()
	set rs=Nothing
End sub

%>

<script language="javascript">
<!--
	function CheckValue()
	{
		var Content,Px,ParentID,ParentID2;
		Content=document.getElementById("Content");
		Px=document.getElementById("Px");
		ParentID=document.getElementById("ParentID");
		ParentID2=document.getElementById("ParentID2");
		
		if(Content.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����дʡ�����ƣ�");
			Content.focus();
			return false;
		}
		
		if((Px.value).search("^-?\\d+(\\.\\d+)?$")!=0)
		{
			alert("����д��Ч�����֣�");
			Px.select();
			return false;
		}
		
		if(ParentID.value=="Null")
		{
			alert("�Բ�������ʡ����Ϣ���������ʡ����Ϣ��");
			return false;
		}
		
		if(ParentID2.value=="Null")
		{
			alert("�Բ�����ѡ������Ϣ��");
			return false;
		}
		return true;
	}

	function Event_Change()
	{
		var ParentID=document.getElementById("ParentID");		
		var xmlhttp=new createxmlhttp();
		queryString="ParentID="+escape(ParentID.value);
		xmlhttp.onreadystatechange =function(){GetBak(xmlhttp);};
		xmlhttp.open("POST","GetCity.asp",true);
		xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
		xmlhttp.send(queryString);
	}
	function GetBak(xmlhttp)
	{
		if(xmlhttp.readyState==4)
		{
			if(xmlhttp.status==200)
			{
				var text=xmlhttp.responseText
				text=text.slice(text.indexOf("$")+1,text.lastIndexOf("$"));
				if(text.indexOf("error")!=-1)
				{
					alert("�Բ��𣬳��ִ�");
					window.location.href="County.asp";
				}
				else
				{
					var ParentID2=document.getElementById("ParentID2");
					if(text.indexOf("������Ϣ")==-1)				
					{
						var Arrays=text.split("|");
						var item_array;
						ParentID2.length=0;
						for(var i=0;i<Arrays.length;i++)
						{
							item_array=Arrays[i].split(",");
							ParentID2.options[i]=new Option(item_array[1],item_array[0])
						}
					}
					else
					{
					
						ParentID2.length=0;
						ParentID2.options[0]=new Option(text,"Null");
					}
					
				}
			}
		}
	}
-->
</script>
