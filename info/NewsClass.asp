<%
Class NewsClass
	Dim Rs,Sql,TuiJian,Page_Size,SortPath,ZiShu,target,ID
	
    Public Sub Set_target(Values)
		target=Values
	End Sub
	Public Sub Set_ID(Values)
		ID=Values
	End Sub
	Private Sub SetRs()
    	set Rs=server.CreateObject("adodb.recordset")
	End Sub
	Public Sub Set_ZiShu(Values)
		ZiShu=Values	
	End Sub
	Public Sub Set_Page_Size(Values)
		Page_Size=Values
	End Sub
	
	Public Sub Set_SortPath(Values)
		SortPath=Values	
	End Sub
	
	Private Sub CloseRs()
		Rs.Close()
		Set Rs=Nothing
	End Sub
	
	Public Sub Set_TuiJian(Values)
		TuiJian=Values
	End Sub
	
	Public Sub TuiJianList()
		SetRs
		Sql="select top "&TuiJian&" ID,NewsName,AddTime from NwebCn_News where ViewFlag=1 and commendFlag=1 and Charindex(SortPath,'"&SortPath&"')>0 order by px desc,AddTime desc"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("<li>")
				response.Write("�����Ƽ���Ϣ�б�")
			response.Write("</li>")
		else
			while not rs.eof
				response.Write("<li>")
					response.Write("<img src='Images/img_20.jpg' align='absmiddle'><a href='NewsView.asp?ID="&rs("ID")&"' style='font-size:12px'>"&StrLeft(rs("NewsName"),26)&"</a>  ")
				response.Write("</li>")
				rs.movenext
			wend
			response.Write("<li style='text-align:right; padding-right:10px;'>")
					response.Write("<a href='News.asp' style='font-size:12px'>���� >> </a>  ")
			response.Write("</li>")
		end if
		CloseRs
	End Sub
	
	Public Sub ScrollTuiJianList()
		TuiJianList
	End sub
	
	Public Sub List()
		SetRs
		Sql="Select ID,NewsName,Source,AddTime,Content from NwebCn_News where ViewFlag=1 and Charindex(SortPath,'"&SortPath&"')>0 order by AddTime desc"
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("<li>")
				response.Write("<img src='images/img_20.jpg' />")
				response.Write("������Ϣ��")
			response.Write("</li>")
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
			
			for i=1 to Page_Size
				if not rs.eof then
					response.write("<li>")
						response.Write("<img src='images/img_20.jpg' />")
						response.Write("<a href='NewsView.asp?ID="&rs("ID")&"' target='"&target&"' style='color:#ff0000'>��Ѷ:")
						response.Write(rs("NewsName"))
						response.Write("</a>")
						response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
						response.Write(rs("Source"))
						response.Write("&nbsp;&nbsp;")		
						response.Write(rs("AddTime"))
						response.Write("<dt>")
							if rs("Content")<>"" then
								response.Write(StrLeft(RemoveHTML(replace(rs("Content"),"../","")),ZiShu))
								response.Write("��<a href='NewsView.asp?ID="&rs("ID")&"' target='"&target&"'>")
								response.Write("<b>����鿴��ϸ����</b>")
								response.Write("</a>��")
							else
								response.Write("��������У����Ժ󡭡�")
							end if
						response.Write("</dt>")
					response.write("</li>")
					rs.movenext
				end if
			next
			if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
		end if
		CloseRS
	End Sub
	
	Private sub Contrl_Page(page,sum_page,total,page_size) 
	dim Url,linkfile,pagewhere,UrlValue
	Url=request.ServerVariables("URL")
	Url=mid(Url,InstrRev(Url,"/")+1)
	linkfile=Url
	Pagewhere="&SortPath="&SortPath
		response.Write("<li style='height:40px;text-align:right;'>")
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
			<%response.Write("</li>")
	end sub
	
	Public Sub Title()
		SetRs
		Sql="Select NewsName from NwebCn_News where id="&ID
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("��������")
		else
			response.Write(Rs("NewsName"))
		end if
		CloseRs
	End Sub
	
	Public Sub Text()
		SetRs
		Sql="Select Content from NwebCn_News where id="&ID
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("��������С���")
		else
			response.Write(Rs("Content"))
		end if
		CloseRs
	End sub
End Class
%>