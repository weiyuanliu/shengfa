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
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>�����м���Ϣ����</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AddCity.asp?Result=Add" onClick='changeAdminFlag("���ʡ����Ϣ")'>����м���Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="City.asp" onClick='changeAdminFlag("�鿴�����м���Ϣ")'>�鿴�����м���Ϣ</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=City" method="post" name="formDel" >
    <tr>
      <td width="120" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID���</strong></font></td>
      <td width="188" align="center"  bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>����ʡ��</strong></font></td>
      <td width="188" height="24" align="center"  bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>������</strong></font></td>
      <td width="183" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>����˳��</strong></font></td>
      <td width="291" align="center"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">���ʱ��</font></strong></td>
      <td width="140" align="center" bgcolor="#8DB5E9"><strong><font color="#FFFFFF">����</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="ȫ" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="��" style="HEIGHT: 18px;WIDTH: 16px;">      </td>
    </tr>
	<%Call CityList(20) %>
  </form>
</table>
</body>
</html>
<%
Sub CityList(Page_Size)
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from City order by px asc,id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='6'>")
				response.Write("������Ϣ��")
			response.Write("</td>")
		response.Write("</tr>")
	else
		rs.pagesize=page_size
		dim sum_page,total,i
		total=rs.recordcount
		sum_page=total \ page_size
		if total mod page_size <>0 then sum_page=sum_page+1
		dim page
		page=trim(request.querystring("page"))
		if page="" or isnull(page) or (not IsNumeric(page)) then
			page=1
		elseif Cint(Page)<=1 then
			page=1
		elseif Cint(page) => sum_page then
			page=sum_page
		else
			page=Cint(page)
		end if
		rs.absolutepage=page
		
		for i=1 to Page_Size 
			if not rs.eof then
				response.Write("<tr bgcolor='#EBF2F9'>")
					
					response.Write("<td align='center'>")
						response.Write(rs("ID"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(GetValue(rs("ParentID"),"Province","Content"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("Content"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("Px"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("AddTime"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write("��<a href='EditCity.asp?ID="&rs("ID")&"'>�� ��</a>��")
						response.Write("<input type='checkbox' name='SelectID' id='SelectID' value='"&rs("ID")&"' style='margin-left:5px;'>ѡ ��")
					response.Write("</td>")
				response.Write("</tr>")
				rs.movenext
			end if
		next
		
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='5'></td>")
			response.Write("<td align='center'>")
				response.Write("<input name='DelRecord' type='submit' value='ɾ ��'>")
			response.Write("</td>")
		response.Write("</tr>")
		
		if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
	end if
	rs.close()
	set rs=Nothing
End sub
%>
<%
sub Contrl_Page(page,sum_page,total,page_size) 
dim Url,linkfile,pagewhere,UrlValue
Url=request.ServerVariables("URL")
Url=mid(Url,InstrRev(Url,"/")+1)
linkfile=Url
UrlValue=""

if UrlValue<>"" then
	pagewhere=UrlValue
end if

	response.Write("<tr bgcolor='#EBF2F9'>")
		response.Write("<td colspan='5' class='Item_list' style='padding-top:5px; padding-bottom:5px; text-align:right;'>")
			response.Write("[���ƣ�"&total&"��] ")
					response.write("[ÿҳ��"&page_size&"��] ")
					response.write("[ҳ�Σ�"&page&"/"&sum_page&"] ")
					if page<=1 then
						response.write("[��ҳ] [��һҳ] ")
					else 
						response.write("<a href='"&linkfile&"?page=1"&pagewhere&"'>")
						response.write("[��ҳ]")
						response.write("</a> ")
						response.write("<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
						response.write("[��һҳ]")
						response.write("</a> ")
					end if
					
					if page < sum_page then
						response.write("<a href='"&linkfile&"?page="&page+1&pagewhere&"'>")
						response.write("[��һҳ]")
						response.write("</a> ")
					else
						response.write("[��һҳ] ")
					end if
					
					if sum_page>1 and page < sum_page then
						response.write("<a href='"&linkfile&"?page="&sum_page&pagewhere&"'>")
						response.write("[ĩҳ]")
						response.write("</a>")
					else
						response.write("[ĩҳ]")
					end if
					dim cc
					response.write(" ת����")%>
					<select name="page" size="1" onchange="javascript:window.location='<%=linkfile%>?page='+this.options[this.selectedIndex].value+'<%=pagewhere%>';">
						<%for cc=1 to sum_page
							if cc=page then
								response.write("<option value='"&cc&"' selected >"&cc&"ҳ")
							else
								response.write("<option value='"&cc&"'>"&cc&"ҳ")
							end if
						next%>
					</select>
		<%response.Write("</td>")
	response.Write("</tr>")
end sub

function  GetValue(ParentID,tablename,Content)
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select "&Content&" from "&tablename&" where id="&ParentID
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		GetValue=rs(Content)		
	end if
	rs.close()
	set rs=Nothing
end function
%>

