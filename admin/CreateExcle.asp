<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<% 
Response.Buffer = True 
Response.ContentType = "application/vnd.ms-excel" 
Response.AddHeader "content-disposition", "inline; filename = "&year(now())&"��"&month(now())&"��"&day(now())&"��"&hour(time)&Minute(time)&Second(time)&".xls"
%> 
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|81,")=0 then 
  response.write ("<script language=javascript> alert('�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�');history.back(-1);</script>")
end if
%>
<%
if Instr(session("AdminPurview"),"|94,")=0 then 
  response.write ("<font color='red')>�㲻���иù���ģ��Ĳ���Ȩ�ޣ��뷵�أ�</font>")
  response.end
end if
'========�ж��Ƿ���й���Ȩ��
%>
<%
	Dim Page,States
	Page=trim(Request.QueryString("Page"))
	States=trim(Request.QueryString("State"))
	Call Asp_Excle(Page,States)
%>
<%
Sub Asp_Excle(Page,States)
	if Page="" or isnull(Page) or not(IsNumeric(Page)) then
		Page=1
	else
		Page=Cint(Page)	
	end if
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	if  States<>""  and States<>"NULL" then
		if Instr(States,"������")>0 then
			sql="select * from NwebCn_Order where HuoDao_FuKuan=1 and (State is null) and fax=0 order by id desc"
		else
			
			sql="select * from NwebCn_Order where charindex(State,'"&States&"')>0 and fax=0 order by id desc"
		end if
	else
		sql="select * from NwebCn_Order where  fax=0 order by id desc"
	end if
 
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("������Ϣ��")
		response.End()
	else
		
		 
		
		
		response.Write("<table border=1 cellpadding=0 cellspacing=1 width='100%'  bordercolor='#cccccc'>")
			response.Write("<tr>")
				
				response.Write("<td align='center'>")
					response.Write("������")
				response.Write("</td>")
				
			 	response.Write("<td align='center'>")
			 		response.Write("֧����ʽ")
			 	response.Write("</td>")
				
				response.Write("<td align='center'>")
					response.Write("��������")
				response.Write("</td>")
				
				response.Write("<td align='center'>")
					response.Write("��ϵ�绰")
				response.Write("</td>")
				
				response.Write("<td align='center'>")
					response.Write("����ʱ��")
				response.Write("</td>")
				
				response.Write("<td align='center'>")
					response.Write("��ַ ����")
				response.Write("</td>")
				
				response.Write("<td align='center'>")
					response.Write("״ ̬")
				response.Write("</td>")
				
			response.Write("</tr>")
			
		while not rs.eof
				 
					response.Write("<tr>")
						response.Write("<td>")
							response.Write(rs("Linkman"))
						response.Write("</td>")
						
					 	response.Write("<td>")
					 	 
						Dim ZiFu_FS
		ZiFu_FS=Split(rs("Remark"),"|")
		Response.Write(ZiFu_FS(1))
		Response.Write(ZiFu_FS(2))
					 	response.Write("</td>")
						
						Response.Write "<td title="&rs("Amount")&" >"&Replace(Replace(replace(Print(rs("Amount")),"�����",""),"һ��0�С�",""),"������0��","")&"</td>" 
						
						response.Write("<td>")
							response.Write(rs("tel"))
						response.Write("</td>")
						
						response.Write("<td>")
							response.Write(rs("AddTime"))
						response.Write("</td>")
						
						response.Write("<td>")
							response.Write(rs("Address")&"("&rs("ZipCode")&")")
						response.Write("</td>")
						
						response.Write("<td>")
							if rs("State")<>"" then
								response.Write(rs("State"))
							else
								if rs("HuoDao_FuKuan") and Not(rs("FuKuan")) then
									response.Write("��������|���ܷ���")
								end if
							end if
						response.Write("</td>")
						
					response.Write("</tr>")
					rs.movenext
				 
			wend
		response.Write("</table>")
	end if
	rs.close()
	set rs=Nothing
End sub
function Print(Amount)
	dim str,i,str1
	str1=""
	str=split(Amount,"|")
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
end function
%>
<script language="javascript">
		<!--
			document.all.WebBrowser.ExecWB(45,1);
		-->
</script>