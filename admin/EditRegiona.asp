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
<script language="javascript" src="../Script/ThreeLd.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="TreeLDClass.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
Dim LianDong,Action,ID
Action=Trim(Request.QueryString("Action"))
ID=Trim(Request.QueryString("ID"))
if id="" or isnull(id) or Not(IsNumeric(id)) then
	Call Message("���ݳ����뷵�أ�")
	response.End()
end if

if Action="EditRecord" then Call EditRecords()
Set LianDong=New LdClass
LianDong.Set_ID(ID)

'����ȫ�ֱ������ڱ������ݿ��е�ֵ 
Dim QY_Names,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_XingZhi
Dim QY_Wai,QY_CaoZuo,QY_BeiZu,QY_AddTime,QY_Px,QY_FanWei
Call FuZhi()
LianDong.Set_QY_ShengFen(QY_ShengFen)
LianDong.Set_QY_City(QY_City)
LianDong.Set_QY_Citys(QY_Citys)
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>������Ϣ���</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AddRegional.asp?Result=Add" onClick='changeAdminFlag("�����Ϣ")'>�����Ϣ</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="Regionallist.asp" onClick='changeAdminFlag("�鿴��Ϣ")'>�鿴��Ϣ</a></td>
  </tr>
  <tr>
    <td height="48" align="center" nowrap  bgcolor="#EBF2F9" style="padding:10px;">
    
    <table width="87%" border="0" cellpadding="4" cellspacing="0">
     <form name="AddRegional" id="AddRegInonal" action="EditRegiona.asp?Action=EditRecord&ID=<%=ID%>" method="post" onSubmit="return Check_AddRegionalValues();">
      <tr>
        <td width="17%" align="right"><strong>�������ƣ�</strong></td>
        <td width="83%"><label>
          <input name="QY_Names" type="text" id="QY_Names" value="<%=QY_Names%>">��
          *����
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>ʡ��ѡ��</strong></td>
        <td><label>
          <select name="QY_ShengFen" id="QY_ShengFen" onChange="ChangEvent('QY_ShengFen','QY_City','QY_Citys','GetLDValue.asp?Action=Two');">
          	<%=LianDong.FirstGread%>
          </select>
        ��*��ѡ</label></td>
      </tr>
      <tr>
        <td align="right"><strong>�м�ѡ��</strong></td>
        <td><label>
          <select name="QY_City" id="QY_City" onChange="ChangEvent('QY_City','QY_Citys','Null','GetLDValue.asp?Action=Three');">
          	<%=LianDong.TwoGread%>
          </select>
        ��*��ѡ</label></td>
      </tr>
      <tr>
        <td align="right"><strong>����ѡ��</strong></td>
        <td><label>
          <select name="QY_Citys" id="QY_Citys">
          	<%=LianDong.ThreeGread%>
          </select>
        ��*��ѡ</label></td>
      </tr>
      <tr>
        <td align="right"><strong>�������ͣ�</strong></td>
        <td><label>
          <input name="QY_Type" type="text" id="QY_Type" value="<%=QY_Type%>">
        ��
          *����
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>�������ʣ�</strong></td>
        <td><label>
          <input name="QY_XingZhi" type="text" id="QY_XingZhi" value="<%=QY_XingZhi%>">
        ��
          *����
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>��������</strong></td>
        <td>
            <INPUT type="hidden" name="QY_FanWei" value="<%=QY_FanWei%>">
            <IFRAME ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=QY_FanWei&style=s_mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME>       
       </td>
      </tr>
      <tr>
        <td align="right"><strong>���������⣺</strong></td>
        <td>
        	<INPUT type="hidden" name="QY_Wai" value="<%=QY_Wai%>">
            <IFRAME ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=QY_Wai&style=s_mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME> 
        </td>
      </tr>
      <tr>
        <td align="right"><strong>�ɲ�����</strong></td>
        <td><label>
          <input type="radio" name="QY_CaoZuo" id="QY_CaoZuo" value="��" <%if QY_CaoZuo then response.Write("checked")%>>
          ��
         �� 
         <input type="radio" name="QY_CaoZuo" id="QY_CaoZuo2" value="��" <%if Not(QY_CaoZuo) then response.Write("checked")%>>
          ��
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>��ע��</strong></td>
        <td>
        	<INPUT type="hidden" name="QY_BeiZu" value="<%=QY_BeiZu%>">
            <IFRAME ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=QY_BeiZu&style=s_mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME> 
        </td>
      </tr>
      <tr>
        <td align="right"><strong>���ʱ�䣺</strong></td>
        <td><label>
          <input name="QY_AddTime" type="text" id="QY_AddTime" value="<%=QY_AddTime%>">
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>����˳��</strong></td>
        <td><label>
          <input name="QY_Px" type="text" id="QY_Px" size="10" value="<%=QY_Px%>">
        ������д����������Ϣ��ֵԽ������Խǰ</label></td>
      </tr>
      <tr>
        <td align="right">&nbsp;</td>
        <td><label>
          <input type="submit" name="tijiao" id="tijiao" value="�� ��" style="margin-left:15px; margin-right:10px;">
          <input type="button" name="GetBak" id="GetBak" value="�� ��" onClick="window.history.go(-1);">
        </label></td>
      </tr>
      </form>
    </table>    </td>    
  </tr>
</table>
<br>
</body>
</html>
<%
Sub Regionallist(Page_Size)
	Dim Rs,Sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select QY_Names,ID,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_AddTime,QY_Px from Regional order by id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='9'>"&vbcrlf)
				response.Write("������Ϣ��")
			response.Write("</td>"&vbcrlf)
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
		
		for i=1 to page_size
			if not rs.eof then
				response.Write("<tr bgcolor='#EBF2F9'>")
					response.Write("<td align='center'>")
						response.Write(rs("ID"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(Get_Values("Province","Content",rs("QY_ShengFen")))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(Get_Values("City","Content",rs("QY_City")))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(Get_Values("County","Content",rs("QY_Citys")))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_Names"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_Type"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_Px"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_AddTime"))
					response.Write("</td>")
				response.Write("</tr>")
				rs.movenext
			end if
		next
		
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='8'></td>")
			response.Write("<td align='center'>")
				response.Write("<input name='DelRecord' type='submit' value='ɾ ��'>")
			response.Write("</td>")
		response.Write("</tr>")
		
		if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
	end if
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
		response.Write("<td colspan='9' class='Item_list' style='padding-top:5px; padding-bottom:5px; text-align:right;'>")
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
					<select name="page" size="1" onChange="javascript:window.location='<%=linkfile%>?page='+this.options[this.selectedIndex].value+'<%=pagewhere%>';">
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

function Get_Values(tablename,Content,ID)
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select "&Content&" from "&tablename&" where id="&ParentID
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		Get_Values=rs(Content)		
	end if
	rs.close()
	set rs=Nothing
end function

Sub EditRecords()
	Dim QY_Names,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_XingZhi
	Dim QY_Wai,QY_CaoZuo,QY_BeiZu,QY_AddTime,QY_Px,QY_FanWei
	
	QY_Names=Trim(Request.Form("QY_Names"))
	QY_ShengFen=Trim(Request.Form("QY_ShengFen"))
	QY_City=Trim(Request.Form("QY_City"))
	QY_Citys=Trim(Request.Form("QY_Citys"))
	QY_Type=Trim(Request.Form("QY_Type"))
	QY_XingZhi=Trim(Request.Form("QY_XingZhi"))
	
	QY_Wai=Trim(Request.Form("QY_Wai"))
	QY_CaoZuo=Trim(Request.Form("QY_CaoZuo"))
	QY_BeiZu=Trim(Request.Form("QY_BeiZu"))
	QY_AddTime=Trim(Request.Form("QY_AddTime"))
	QY_Px=Trim(Request.Form("QY_Px"))
	QY_FanWei=Trim(Request.Form("QY_FanWei"))
	
	if QY_Names="" or isnull(QY_Names) then
		Call Message("����д���֣�")
		response.End()	
	end if
	
	if QY_ShengFen="" or isnull(QY_ShengFen) or QY_ShengFen="Null" or not(IsNumeric(QY_ShengFen)) then
		Call Message("���ݲ���Ϊ�գ��뷵�أ�")	
		response.End()
	end if
	
	if QY_City="" or isnull(QY_City) or QY_City="Null" or not(IsNumeric(QY_City)) then
		Call Message("���ݲ���Ϊ�գ��뷵�أ�")	
		response.End()
	end if
	
	if QY_Citys="" or isnull(QY_Citys) or QY_Citys="Null" or not(IsNumeric(QY_Citys)) then
		Call Message("���ݲ���Ϊ�գ��뷵�أ�")	
		response.End()
	end if
	
	if QY_Type="" or isnull(QY_Type) then
		Call Message("���ݲ���Ϊ�գ��뷵�أ�")	
		response.End()
	end if
	
	if QY_XingZhi="" or isnull(QY_XingZhi) then
		Call Message("���ݲ���Ϊ�գ��뷵�أ�")	
		response.End()
	end if
	
	if QY_Px="" or isnull(QY_Px) or Not(IsNumeric(QY_Px)) then
		Call Message("���ݳ����뷵�أ�")	
		response.End()
	end if
	
	if QY_FanWei="" or isnull(QY_FanWei) then
		Call Message("���ݲ���Ϊ�գ��뷵�أ�")	
		response.End()
	end if
	
	Dim Rs,Sql
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select * from Regional where QY_Names='"&QY_Names&"' and QY_ShengFen="&QY_ShengFen&" and QY_City="&QY_City&" and QY_Citys="&QY_Citys&" and id not in("&ID&")"
	Rs.open Sql,conn,1,1
	if Not rs.eof and Not rs.bof then
		rs.close()
		set rs=Nothing
		Call Message("�Բ��𣬴˼�¼�Ѿ����ڣ��뷵�أ�")
		response.End()
		exit sub
	end if
	rs.close
	Sql="Select top 1 * from Regional Where ID="&ID
	Rs.open sql,conn,1,3
		rs("QY_Names")=QY_Names
		rs("QY_ShengFen")=QY_ShengFen
		rs("QY_City")=QY_City
		rs("QY_Citys")=QY_Citys
		rs("QY_Type")=QY_Type
		rs("QY_XingZhi")=QY_XingZhi
		rs("QY_FanWei")=QY_FanWei
		rs("QY_Wai")=QY_Wai
		if QY_CaoZuo then
			rs("QY_CaoZuo")=true
		else
			rs("QY_CaoZuo")=false
		end if
		rs("QY_BeiZu")=QY_BeiZu
		if QY_AddTime<>"" then
			rs("QY_AddTime")=QY_AddTime
		else
			rs("QY_AddTime")=Now()
		end if	
		rs("QY_Px")=QY_Px
	rs.update()
	rs.close()
	set rs=Nothing
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('��¼�޸ĳɹ���');"&vbcrlf)
		response.Write("window.location.href='Regionallist.asp';")
	response.Write("</script>"&vbcrlf)
End Sub

Sub Message(str)
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('"&str&"');")&vbcrlf
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>"&vbcrlf)
End Sub

Sub FuZhi() '���ڶ�ȡ���ݿ��ĳ����¼��ֵ����������ȫ�ֱ�����
	Dim Rs,Sql
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * from Regional where id="&ID
	Rs.Open Sql,conn,1,1
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		Call Message("�Բ��𣬼�¼δ�ҵ���")
		response.End()
		exit sub
	else
		QY_Names=Rs("QY_Names")
		QY_ShengFen=Rs("QY_ShengFen")
		QY_City=Rs("QY_City")
		QY_Citys=rs("QY_Citys")
		QY_Type=rs("QY_Type")
		QY_XingZhi=rs("QY_XingZhi")
		
		QY_Wai=rs("QY_Wai")
		QY_CaoZuo=rs("QY_CaoZuo")
		QY_BeiZu=rs("QY_BeiZu")
		QY_AddTime=rs("QY_AddTime")
		QY_Px=rs("QY_Px")
		QY_FanWei=rs("QY_FanWei")
	end if
	rs.close()
	set rs=Nothing
End Sub
%>

