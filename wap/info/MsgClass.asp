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
		Sql="Select top "&HomeNumbers&" MesName,Content,Mobile,Linkman,AddTime,ReplyContent from NwebCn_Message where ViewFlag=1 and SecretFlag=1 order by AddTime desc"
		rs.open sql,conn,1,1
		response.Write("<ul>")
		if rs.eof and rs.bof then
			response.Write("<li>")
				response.Write("�Բ���������Ϣ��")
			response.Write("</li>")
		else
			while not rs.eof
					response.Write("<li><span>��&nbsp;&nbsp;�ԣ�</span><font style=""width:300px;"">"&Rs("linkman")&" ������</font><label>IP��"&rs("Mobile")&"&nbsp; &nbsp; ���ڣ�"&rs("AddTime")&"</label></li>")
					response.Write("<li><span>��&nbsp;&nbsp;�⣺</span><font>"&rs("MesName")&"</font></li>")
					response.Write("<li><span>��&nbsp;&nbsp;�ԣ�</span><font>"&ReStrReplace(rs("Content"))&"</font></li>")
					response.Write("<li class=""reds""><span>��&nbsp;&nbsp;����</span><font style=""width:720px;"">���ã���л��������!&nbsp;"&ReStrReplace(rs("ReplyContent"))&"</font></li>")
				rs.movenext
			wend
		end if
		response.Write("</ul>")
		CloseRs
	End Sub
	
	Public Sub List()
		SetRs
		Sql="Select ID,MesName,Content,Linkman,Mobile,AddTime,ReplyContent,ReplyTime from NwebCn_Message where ViewFlag=1 order by AddTime desc"
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("����������Ϣ��")
		else
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
								ViewMsgText rs("Linkman"),rs("Mobile"),rs("MesName"),rs("Content"),rs("ReplyContent"),rs("ReplyTime"),rs("AddTime")
						rs.movenext
					end if
				next
				if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size)
		end if
		CloseRs
	End Sub
	
	Private Sub ViewMsgText(Linkman,Mobile,MesName,Conten,ReplyContent,ReplyTime,AddTime)
					response.Write("<ul>")
					response.Write("<li><span>��&nbsp;&nbsp;�ԣ�</span><font style=""width:300px;"">")
					response.Write(Rs("linkman"))
					response.Write("������</font><label>IP��")
					response.Write(Mobile)
					response.Write("&nbsp; &nbsp; ���ڣ�"&AddTime&"</label></li>")
					response.Write("<li><span>��&nbsp;&nbsp;�⣺</span><font>")
					response.Write(ReStrReplace(MesName)&"</font></li>")
					response.Write("<li><span>��&nbsp;&nbsp;�ԣ�</span><font style=""max-width:460px;"">")
					response.Write(ReStrReplace(Conten)&"</font></li>")
			if ReplyContent<>"" then
					response.Write("<li class=""reds""><span>��&nbsp;&nbsp;����</span><font style=""max-width:460px;"">")
					response.Write("���ã���л��������!&nbsp;"&ReStrReplace(ReplyContent)&"</font></li>")
			end if
					response.Write("</ul>")
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
	
		response.Write("<div class=""page"">")
			response.Write("<DIV class=""pagelistbox"">")
				response.Write("<SPAN>���ƣ�"&total&"��")
						response.write(" / ÿҳ��"&page_size&"��")
						response.write(" / ҳ�Σ�"&page&"/"&sum_page&"</SPAN>")
						response.write("<br/>")
						if page<=1 then
							response.write("<span class=""current"">��ҳ</span><span class=""current"">��һҳ</span>")
						else 
							response.write("<a href='"&linkfile&"?page=1"&pagewhere&"'>")
							response.write("��ҳ")
							response.write("</a>")
							response.write("<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
							response.write("��һҳ")
							response.write("</a>")
						end if
						
						if page < sum_page then
							response.write("<a href='"&linkfile&"?page="&page+1&pagewhere&"'>")
							response.write("��һҳ")
							response.write("</a>")
						else
							response.write("<span class=""current"">��һҳ</span>")
						end if
						
						if sum_page>1 and page < sum_page then
							response.write("<a href='"&linkfile&"?page="&sum_page&pagewhere&"'>")
							response.write("ĩҳ")
							response.write("</a>")
						else
							response.write("<span class=""current"">ĩҳ</span>")
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
			<%response.Write("</div>")
		response.Write("</div>")
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
		 
		Sql="Select * from MsgData where  datalength(replay)>0 Order by ReplayTime desc"
		 
		Rs.open Sql,Conn,1,1
		IF Rs.eof and Rs.bof then
			Response.Write("�Բ���������������Ϣ��")
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
						if i<10 and page=1 then
							response.Write("<div style='margin:0px; padding:0px; font-size:14px;'><span style='float:left;'><font color='ff0000'>"&rs("Msg_Name")&"</font>�������ã�</span>�����Ϣ�ѷ���,�����ϱ����붩��ʱ�ĵ绰�Ų鿴��<span style='float:right;'></span></div><div style='clear:both; margin:0px; margin-top:-10px; padding:0px;'>�ظ�ʱ��:"&Rs("ReplayTime") &"<img src='images/news.gif' style='width:auto;' /></div>")
						else
							response.Write("<div style='margin:0px; padding:0px; font-size:14px;'><span style='float:left;'><font color='ff0000'>"&rs("Msg_Name")&"</font>�������ã�</span>�����Ϣ�ѷ���,�����ϱ����붩��ʱ�ĵ绰�Ų鿴��<span style='float:right;'></span></div>")
							'response.Write("<span style='float:left; font-size:14px;'><font color='ff0000' >"&rs("Msg_Name")&"</font>�������ã�</span><span style='float:right;'>�����Ϣ�ѷ���,�����ϱ����붩��ʱ�ĵ绰�Ų鿴��</span>")
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
			Response.Write("�Բ���������������Ϣ��")
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
					response.Write(rs("Msg_Name")&"�������ã�<br />")
						if rs("Replay")="" or isnull(rs("Replay")) then
							response.Write("<span style='margin-left:20px;'>")
								response.Write("�Բ������������Ϣ������Ա��δ�ظ���")
							response.Write("</span>")						
						else
							response.Write("<span style='margin-left:20px;'>")
								response.Write(rs("Replay"))
							response.Write("</span>")		
							response.Write("<br />")
							response.Write("<span style='margin-top:10px; text-align:right; margin-right:10px; float:right;'>")
								response.Write("�ظ�ʱ�䣺"&FormatDate(rs("ReplayTime"),4))
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
			response.Write("<div style='text-align:left;'>")
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
									response.write("<a href='"&linkfile&"?page="&page+1&pagewhere&"' style='color:#000000'>")
									response.write("[��һҳ]")
									response.write("</a> ")
								else
									response.write("[��һҳ] ")
								end if
								
								if sum_page>1 and page < sum_page then
									response.write("<a href='"&linkfile&"?page="&sum_page&pagewhere&"' style='color:#000000'>")
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
%>