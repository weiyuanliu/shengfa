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
						if i<10 and page=1 then
							response.Write("<li><span>"&left(rs("Msg_Name"),1)&"先生/小姐</span>&nbsp;&nbsp;您好，你的订单已经发货，请注意查收！</li>")
						else
							response.Write("<li><span>"&left(rs("Msg_Name"),1)&"先生/小姐</span>&nbsp;&nbsp;您好，你的订单已经发货，请注意查收！</li>")
						end if
					rs.movenext
				end if
			next
		End IF
		CloseRs
	End Sub
End Class
%>