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
				response.Write("暂无推荐信息列表！")
			response.Write("</li>")
		else
			while not rs.eof
				response.Write("<li>")
					response.Write("<img src='Images/img_20.jpg' align='absmiddle'><a href='NewsView.asp?ID="&rs("ID")&"' style='font-size:12px'>"&StrLeft(rs("NewsName"),26)&"</a>  ")
				response.Write("</li>")
				rs.movenext
			wend
			response.Write("<li style='text-align:right; padding-right:10px;'>")
					response.Write("<a href='News.asp' style='font-size:12px'>更多 >> </a>  ")
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
				response.Write("暂无信息！")
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
						response.Write("<a href='NewsView.asp?ID="&rs("ID")&"' target='"&target&"' style='color:#ff0000'>快讯:")
						response.Write(rs("NewsName"))
						response.Write("</a>")
						response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
						response.Write(rs("Source"))
						response.Write("&nbsp;&nbsp;")		
						response.Write(rs("AddTime"))
						response.Write("<dt>")
							if rs("Content")<>"" then
								response.Write(StrLeft(RemoveHTML(replace(rs("Content"),"../","")),ZiShu))
								response.Write("（<a href='NewsView.asp?ID="&rs("ID")&"' target='"&target&"'>")
								response.Write("<b>点击查看详细内容</b>")
								response.Write("</a>）")
							else
								response.Write("内容添加中，请稍后……")
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
				response.Write("[共计："&total&"条] ")
						response.write("[每页："&page_size&"条] ")
						response.write("[页次："&page&"/"&sum_page&"] ")
						if page<=1 then
							response.write("[首页]　[上一页] ")
						else 
							response.write("[<a href='"&linkfile&"?page=1"&pagewhere&"'>")
							response.write("首页")
							response.write("</a>] ")
							response.write("[<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
							response.write("上一页")
							response.write("</a>] ")
						end if
						
						if page < sum_page then
							response.write("[<a href='"&linkfile&"?page="&page+1&pagewhere&"'>")
							response.write("下一页")
							response.write("</a>]　")
						else
							response.write("[下一页] ")
						end if
						
						if sum_page>1 and page < sum_page then
							response.write("[<a href='"&linkfile&"?page="&sum_page&pagewhere&"'>")
							response.write("末页")
							response.write("</a>]")
						else
							response.write("[末页]")
						end if
						dim cc
						response.write(" 转到：")%>
						<select name="page" size="1" onchange="javascript:window.location='<%=linkfile%>?page='+this.options[this.selectedIndex].value+'<%=pagewhere%>';">
							<%for cc=1 to sum_page
								if cc=page then
									response.write("<option value='"&cc&"' selected >"&cc&"页")
								else
									response.write("<option value='"&cc&"'>"&cc&"页")
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
			response.Write("新闻中心")
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
			response.Write("内容添加中……")
		else
			response.Write(Rs("Content"))
		end if
		CloseRs
	End sub
End Class
%>