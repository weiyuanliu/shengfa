<%
Class MsgClass
	Dim Rs,Sql,HomeNumbers,Page_Size
	
	Public Sub Set_Page_Size(Values)
		Page_Size=Values
	End Sub
	Private Sub SetRs()
		Set Rs=Server.CreateObject("Adodb.Recordset")
	End Sub
	Private Sub CloseRs()
		Rs.Close()
		Set Rs=Nothing
	End Sub
	
	Public Sub Set_HomeNumbers(Values)
		HomeNumbers=Values
	End Sub
	
	Public Sub TuiJianList()
		SetRs
		Sql="Select top "&HomeNumbers&" MesName,Content,Mobile,Linkman,AddTime from NwebCn_Message where ViewFlag and SecretFlag order by AddTime desc"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("<li>")
				response.Write("对不起，暂无信息！")
			response.Write("</li>")
		else
			while not rs.eof
				response.Write("<li>")
					response.Write("[来自 "&Rs("linkman")&" 的朋友@"& rs("addtime")&"]<br/>")
					response.Write("IP地址："&rs("Mobile")&"<br/>")
					response.Write("标题 : "&rs("MesName")&"<br/>")
					response.Write("评论留言 :"&ReStrReplace(rs("Content")))
				response.Write("</li>")
				rs.movenext
			wend
		end if
		CloseRs
	End Sub
	
	Public Sub List()
		SetRs
		Sql="Select ID,MesName,Content,Linkman,Mobile,AddTime,ReplyContent,ReplyTime from NwebCn_Message where ViewFlag order by AddTime desc"
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("暂无留言信息！")
		else
			response.Write("<table border=0 cellpadding=0 cellspacing=0 wdith='100%'>")
				rs.pagesize=Page_Size
				dim sum_page,total,i
				total=rs.recordcount
				sum_page=total \ Page_Size
				if total mod Page_Size <>0 then sum_page=sum_page+1
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
						response.Write("<tr>")
							response.Write("<td>")
								ViewMsgText rs("Linkman"),rs("Mobile"),rs("MesName"),rs("Content"),rs("ReplyContent"),rs("ReplyTime"),rs("AddTime")
							response.Write("</td>")
						response.Write("</tr>")
						response.Write("<tr><td height=10></td></tr>")
						response.Write("<tr><td Class='Msg_Line'></tr>")
						response.Write("<tr><td height=10></td></tr>")
						rs.movenext
					end if
				next
				if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
			response.Write("</table>")
		end if
		CloseRs
	End Sub
	
	Private Sub ViewMsgText(Linkman,Mobile,MesName,Conten,ReplyContent,ReplyTime,AddTime)
		response.Write("<table border=0 cellpadding=0 cellspacing=0 width='100%'>")
			response.Write("<tr>")
				response.Write("<td>")
					response.Write("[来自")
					response.Write("&nbsp;&nbsp;")
					response.Write(Rs("linkman"))
					response.Write("&nbsp;&nbsp;")
					response.Write("的朋友@"& AddTime&"]")
				response.Write("</td>")
			response.Write("</tr>")
			response.Write("<tr>")
				response.Write("<td>")
					response.Write("IP地址：")
					response.Write(Mobile)
				response.Write("</td>")
			response.Write("</tr>")
			response.Write("<tr>")
				response.Write("<td>")
					response.Write("标题：")
					response.Write(ReStrReplace(MesName))
				response.Write("</td>")
			response.Write("</tr>")
			response.Write("<tr>")
				response.Write("<td>")
					response.Write("内容：")
					response.Write(ReStrReplace(Conten))
				response.Write("</td>")
			response.Write("</tr>")
			if ReplyContent<>"" then
				response.Write("<tr>")
					response.Write("<td>")
						response.Write("网站回复：")
						response.Write(ReStrReplace(ReplyContent))
					response.Write("</td>")
				response.Write("</tr>")
			end if
		response.Write("</table>")
	End Sub
	
	Private Sub Contrl_Page(page,sum_page,total,page_size) 
	dim Url,linkfile,pagewhere,UrlValue
	Url=request.ServerVariables("URL")
	Url=mid(Url,InstrRev(Url,"/")+1)
	linkfile=Url
	UrlValue=""
	if trim(Request("btype_id"))<>"" and IsNumeric(trim(Request("btype_id"))) then
		UrlValue=UrlValue&"&btype_id="&trim(Request("btype_id"))
	end if
	
	if Trim(Request("stype_id"))<>"" and IsNumeric(Trim(Request("stype_id"))) then
		UrlValue=UrlValue&"&stype_id="&Trim(Request("stype_id"))
	end if
	
	if Trim(Request("Action"))<>"" then
		UrlValue=UrlValue&"&Action="&Trim(Request("Action"))
	end if
	
	if UrlValue<>"" then
		Pagewhere=UrlValue
	end if
	
		response.Write("<tr>")
			response.Write("<td class='page' align='right'>")
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
			<%response.Write("</td>")
		response.Write("</tr>")
	end sub

End Class
%>

<%
Class ViewClass
	Dim TelNumber,Rs,Sql,Page_Size
	
	Private Sub SetRs()
		Set Rs=Server.CreateObject("Adodb.Recordset")
	End Sub
	
	Public Sub Set_Page_Size(Values)
		Page_Size=Values
		 
	End Sub
	
	Private Sub CloseRs()
		Rs.close()
		Set Rs=Nothing	
	End Sub
	
	Public Sub Set_TelNumber(Values)
		TelNumber=Values
	End Sub
	
	Public Function IsTrue()
		SetRs
		Sql="Select id from MsgData where Msg_TelPhone='"&TelNumber&"'"
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			IsTrue=false
		else
			IsTrue=true
		end if
		CloseRs
	End Function
	
	Public Sub ViewList()
		SetRs	
		 
		Sql="Select * from MsgData where  replay<>'' Order by ReplayTime desc"
		 
		Rs.open Sql,Conn,1,1
		IF Rs.eof and Rs.bof then
			Response.Write("对不起，你尚无留言信息！")
		Else
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
			
			For i=1 to page_size
				if not rs.eof then
					Response.Write("<div style='clear:both; line-height:25px;'>")
						if i<11 and page=1 then
							response.Write("<div style='margin:0px; padding:0px;'><span style='float:left;'>"&rs("Msg_Name")&"先生您好！</span><span style='float:right;'><font color='ff0000' >你的信息已反馈,请在上边输入正确的电话号才有权查看</font></span></div><div style='clear:both; margin:0px; margin-top:-10px; padding:0px;'>回复时间:"&Rs("ReplayTime") &"<img src='images/news.gif' /></div>")
						else
							response.Write("<span style='float:left;'>"&rs("Msg_Name")&"先生您好！</span><span style='float:right;'><font color='ff0000' >你的信息已反馈,请在上边输入正确的电话号才有权查看</font></span>")
						end if 
					Response.Write("</div>")
					rs.movenext
				end if
			next
			if sum_page>1 then 
				Response.Write("<div style='clear:both;'>")
				call Contrl_Page(page,sum_page,total,page_size) 
				response.Write("</div>")
			end if
		End IF
		CloseRs
	End Sub
	Public Sub ViewContent()
		SetRs	
		 
		Sql="Select * from MsgData where Msg_TelPhone='"&TelNumber&"' Order by Msg_Time desc,Id desc"
		 
		Rs.open Sql,Conn,1,1
		IF Rs.eof and Rs.bof then
			Response.Write("对不起，你尚无留言信息！")
		Else
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
			
			For i=1 to page_size
				if not rs.eof then
				response.Write("<div style='border:#B6DAEA 1px solid; background:#FFFFFF; width:90%; line-height:20px; color:#030303; padding:10px; margin-bottom:5px; text-align:left;'>")
					response.Write(rs("Msg_Name")&"先生您好！<br />")
						if rs("Replay")="" or isnull(rs("Replay")) then
							response.Write("<span style='margin-left:20px;'>")
								response.Write("对不起，你的留言信息，管理员尚未回复！")
							response.Write("</span>")						
						else
							response.Write("<span style='margin-left:20px;'>")
								response.Write(rs("Replay"))
							response.Write("</span>")		
							response.Write("<br />")
							response.Write("<span style='margin-top:10px; text-align:right; margin-right:10px; float:right;'>")
								response.Write("回复时间："&FormatDate(rs("ReplayTime"),4))
							response.Write("</span>")
						end if
				response.Write("</div>")
				rs.movenext
				end if
			next
			if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
		End IF
		CloseRs
	End Sub
	
	Private Sub Contrl_Page(page,sum_page,total,page_size) 
		dim Url,linkfile,pagewhere,UrlValue
		Url=request.ServerVariables("URL")
		Url=mid(Url,InstrRev(Url,"/")+1)
		linkfile=Url
		UrlValue=""

		if TelNumber<>"" then
			if UrlValue<>"" then
				UrlValue=UrlValue&"&TelPhone="&server.URLEncode(Trim(TelNumber))
			else
				UrlValue=UrlValue&"&TelPhone="&server.URLEncode(Trim(TelNumber))
			end if
		end if
		
		if Trim(Request("Action"))<>"" then
			if UrlValue<>"" then
				UrlValue=UrlValue&"&Action="&server.URLEncode(Trim(Request("Action")))
			else
				UrlValue=UrlValue&"&Action="&server.URLEncode(Trim(Request("Action")))
			end if
		end if
		
		pagewhere=UrlValue
			response.Write("<div style='text-align:right;'>")
						response.Write("[共计："&total&"条] ")
								response.write("[每页："&page_size&"条] ")
								response.write("[页次："&page&"/"&sum_page&"] ")
								if page<=1 then
									response.write("[首页] [上一页] ")
								else 
									response.write("<a href='"&linkfile&"?page=1"&pagewhere&"'>")
									response.write("[首页]")
									response.write("</a> ")
									response.write("<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
									response.write("[上一页]")
									response.write("</a> ")
								end if
								
								if page < sum_page then
									response.write("<a href='"&linkfile&"?page="&page+1&pagewhere&"' style='color:#000000'>")
									response.write("[下一页]")
									response.write("</a> ")
								else
									response.write("[下一页] ")
								end if
								
								if sum_page>1 and page < sum_page then
									response.write("<a href='"&linkfile&"?page="&sum_page&pagewhere&"' style='color:#000000'>")
									response.write("[末页]")
									response.write("</a>")
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
					<%response.Write("</div>")
			End Sub
End Class
Public Function LookAdd(Sip)
  Dim Str1,Str2,Str3,Str4
  Dim Num
  Dim Irs
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
    adb = "DATAbase/ip.mdb"
    aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(adb)
    Set AConn = Server.CreateObject("ADODB.Connection")
    aConn.Open aConnStr
    sql="select country from IPTABLE where StartIPnum <="&num&" and EndIPnum >="&num
    Set irs=AConn.Execute(sql)
    If irs.eof And irs.bof Then 
     LookAdd="中国"
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
%>