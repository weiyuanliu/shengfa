<%
Class SearchQY
	Dim Rs,Sql,Page_Size
	Private Sub SetRs()
		Set Rs=server.CreateObject("adodb.recordset")
	End Sub
	
	Private Sub CloseRs()
		Rs.Close()
		Set Rs=Nothing	
	End Sub
	
	Public Sub Set_Page_Size(Values)
		Page_Size=Values	
	End Sub
	
	Public Sub SearcList()
		response.Write("<table border=0 cellpadding=5 cellspacing=1 width='100%' bgcolor='#484848'>")
			SearchTitle
			List
		response.Write("</table>")
	End sub
	
	Private Sub List()
		Dim QY_ShengFen,QY_City,QY_Citys,DataWhere
		QY_ShengFen=Trim(Request("QY_ShengFen"))
		QY_City=Trim(Request("QY_City"))
		QY_Citys=Trim(Request("QY_Citys"))
		SetRs
		if QY_ShengFen<>"" and Not(isnull(QY_ShengFen)) and QY_ShengFen<>"Null" then
			if DataWhere="" or isnull(DataWhere) then
				DataWhere="where QY_ShengFen="&QY_ShengFen
			else
				DataWhere=DataWhere&" and QY_ShengFen="&QY_ShengFen
			end if
		end if
		
		if QY_City<>"" and Not(isnull(QY_City)) and QY_City<>"Null" then
			if DataWhere="" or isnull(DataWhere) then
				DataWhere="Where QY_City="&QY_City
			else
				DataWhere=DataWhere&" and QY_City="&QY_City
			end if
		end if
		
		if QY_Citys<>"" and Not(isnull(QY_Citys)) and QY_Citys<>"Null" then
			if DataWhere="" or isnull(DataWhere) then
				DataWhere="Where QY_Citys="&QY_Citys
			else
				DataWhere=DataWhere&" and QY_Citys="&QY_Citys
			end if
		end if
		
		Sql="Select * from Regional "&DataWhere&" Order by QY_Px desc,id asc"
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("<tr bgcolor='#1B1B1B'>")		
				response.Write("<td colspan='10'>")
					response.Write("�Բ���������Ҫ���ҵ���Ϣ��")
				response.Write("</td>")
			response.Write("</tr>")
		else
			rs.pagesize=Page_Size
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
			
			For i=1 to Page_Size
				if not rs.eof then
					ViewList rs("QY_Names"),rs("QY_ShengFen"),rs("QY_City"),rs("QY_Citys"),rs("QY_Type"),rs("QY_XingZhi"),rs("QY_FanWei"),rs("QY_Wai"),rs("QY_CaoZuo"),rs("QY_BeiZu")
					rs.movenext
				end if
			next
			if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
		end if
		CloseRs
	End Sub
	
	Private Sub SearchTitle()
		response.Write("<tr bgcolor='#303030' height='25'>")
			response.Write("<td align='center' width='8%'>")
				response.Write("<strong>��������</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='8%'>")
				response.Write("<strong>ʡ����</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='8%'>")
				response.Write("<strong>��</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='10%'>")
				response.Write("<strong>��/��/��</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='7%'>")
				response.Write("<strong>��������</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='7%'>")
				response.Write("<strong>��������</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='15%'>")
				response.Write("<strong>լ���ͷ�������</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='15%'>")
				response.Write("<strong>լ���Ͳ��ɷ�������</strong>")
			response.Write("</td>")
			
			response.Write("<td align='center' width='5%'>")
				response.Write("<strong>����</strong>")
			response.Write("</td>")

			response.Write("<td align='center' width='15%'>")
				response.Write("<strong>��ע</strong>")
			response.Write("</td>")
			
		response.Write("</tr>")
	End Sub
	
	'//������ʾ���б�����
	Private Sub ViewList(QY_Names,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_XingZhi,QY_FanWei,QY_Wai,QY_CaoZuo,QY_BeiZu)
		response.Write("<tr bgcolor='#1B1B1B'>")
			response.Write("<td align='center'>")
				response.Write(QY_Names)
			response.Write("</td>")
			
			response.Write("<td align='center'>")
				response.Write(Get_Values("Province","Content",QY_ShengFen))
			response.Write("</td>")
			
			response.Write("<td align='center'>")
				response.Write(Get_Values("City","Content",QY_City))
			response.Write("</td>")
			
			response.Write("<td align='center'>")
				response.Write(Get_Values("County","Content",QY_Citys))
			response.Write("</td>")
			
			response.Write("<td align='center'>")
				response.Write(QY_Type)
			response.Write("</td>")
			
			response.Write("<td align='center'>")
				response.Write(QY_XingZhi)
			response.Write("</td>")
			
			response.Write("<td>")
				response.Write(QY_FanWei)
			response.Write("</td>")
			
			response.Write("<td>")
				response.Write(QY_Wai)
			response.Write("</td>")
			
			response.Write("<td align='center'>")
				if QY_CaoZuo then 
					response.Write("��")
				else
					response.Write("��")
				end if
			response.Write("</td>")
			
			response.Write("<td>")
				response.Write(QY_BeiZu)
			response.Write("</td>")
		response.Write("</tr>")
	End Sub

	Private Function Get_Values(TableName,ZiDuan,ID)
		Dim Rs,Sql
		Set Rs=server.CreateObject("Adodb.Recordset")
		Sql="Select "&ZiDuan&" from "&TableName&" Where ID="&ID
		Rs.Open Sql,conn,1,1
		if rs.eof and rs.bof then
			Get_Values="�Բ���������Ϣ��"
		else
			Get_Values=Rs(ZiDuan)
		end if
		Rs.close()
		Set Rs=Nothing
	End Function
	
	'//��ҳ������
	
	Private sub Contrl_Page(page,sum_page,total,page_size) 
	dim Url,linkfile,pagewhere,UrlValue
	Url=request.ServerVariables("URL")
	Url=mid(Url,InstrRev(Url,"/")+1)
	linkfile=Url
	Pagewhere="&"
		
	if Trim(Request("QY_ShengFen"))<>"" and IsNumeric(Trim(Request("QY_ShengFen"))) then
		Pagewhere=Pagewhere&"QY_ShengFen="&Trim(Request("QY_ShengFen"))
	end if	
	
	if Trim(Request("QY_City"))<>"" and IsNumeric(Trim(Request("QY_City"))) then
		Pagewhere=Pagewhere&"&QY_City="&Trim(Request("QY_City"))
	end if	
	
	if Trim(Request("QY_Citys"))<>"" and IsNumeric(Trim(Request("QY_Citys"))) then
		Pagewhere=Pagewhere&"&QY_Citys="&Trim(Request("QY_Citys"))
	end if
	
		response.Write("<tr bgcolor='#1B1B1B'>")
		response.Write("<td colspan='10' style='height:20px;text-align:right;'>")
				response.Write("[���ƣ�"&total&"��] ")
						response.write("[ÿҳ��"&page_size&"��] ")
						response.write("[ҳ�Σ�"&page&"/"&sum_page&"] ")
						if page<=1 then
							response.write("[��ҳ]��[��һҳ] ")
						else 
							response.write("[<a href='"&linkfile&"?page=1"&pagewhere&"'>")
							response.write("��ҳ")
							response.write("</a>] ")
							response.write("[<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
							response.write("��һҳ")
							response.write("</a>] ")
						end if
						
						if page < sum_page then
							response.write("[<a href='"&linkfile&"?page="&page+1&pagewhere&"'>")
							response.write("��һҳ")
							response.write("</a>]��")
						else
							response.write("[��һҳ] ")
						end if
						
						if sum_page>1 and page < sum_page then
							response.write("[<a href='"&linkfile&"?page="&sum_page&pagewhere&"'>")
							response.write("ĩҳ")
							response.write("</a>]")
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
End Class
%>