<%
Class LdClass
	Dim Rs,Sql,FirstID,TowID,ID,QY_ShengFen,QY_City,QY_Citys
	Dim Request_ShengFen,Request_City,Request_Citys
	
    Private Sub SetRs()
		Set Rs=Server.CreateObject("adodb.Recordset")
	End Sub
	
	Private Sub CloseRs()
		Rs.Close()
		Set Rs=Nothing
	End Sub
	
	Public Sub Set_Request_ShengFen(Values)
		Request_ShengFen=values	
	end sub
	
	Public Sub Set_Request_City(Values)
		Request_City=Values	
	end sub
	
	Public Sub Set_Request_Citys(Values)
		Request_Citys=Values	
	end sub
	
	Public Sub Set_QY_ShengFen(Values)
		QY_ShengFen=values
	End Sub
	
	Public Sub Set_QY_City(Values)
		QY_City=Values
	End Sub
	
	Public Sub Set_QY_Citys(Values)
		QY_Citys=Values
	End Sub
	
	Public Sub Set_FirstID(Values)
		FirstID=Values
	End Sub
	
	Public Sub Set_ID(Values)
		ID=Values
	End Sub
	
	Public Sub FirstGread()
		SetRs
		Sql="Select * From Province order by Px desc,ID asc"
		Rs.open Sql,Conn,1,1
		if Rs.eof And Rs.bof then
			response.Write("<option value='Null'>暂无信息！</option>")
		else
			FirstID=Rs("ID")
			While not rs.eof
				if QY_ShengFen<>"" and IsNumeric(QY_ShengFen) and QY_ShengFen=rs("ID") then
					response.Write("<option value='"&Rs("ID")&"' selected>"&rs("Content")&"</option>")
					FirstID=Rs("ID")
				else
					if Request_ShengFen<>"" and IsNumeric(Request_ShengFen) then
						if Cint(Request_ShengFen)=rs("ID") then
							response.Write("<option value='"&Rs("ID")&"' selected>"&rs("Content")&"</option>")
							FirstID=Rs("ID")
						else
							response.Write("<option value='"&Rs("ID")&"'>"&rs("Content")&"</option>")
						end if
					else
						response.Write("<option value='"&Rs("ID")&"'>"&rs("Content")&"</option>")
					end if
					
				end if
				rs.movenext
			wend
		end if
		CloseRs
	End Sub
	
	Public Sub TwoGread()
		if FirstID<>"Null" and FirstID<>"" and IsNumeric(FirstID) then
			SetRs
			Sql="Select * from City where ParentID="&FirstID&" order by px desc,id asc"
			Rs.Open Sql,Conn,1,1
			if Rs.eof And Rs.bof then
				TowID="Null"
				response.Write("<option value='Null'>暂无信息！</option>")
			else
				TowID=rs("ID")
				while not rs.eof
					if QY_City<>"" and IsNumeric(QY_City) and QY_City=rs("ID") then
						response.Write("<option value='"&rs("ID")&"' selected>"&rs("Content")&"</option>")	
						TowID=rs("ID")			
					else
						if Request_City<>"" and IsNumeric(Request_City) then
							if Cint(Request_City)=rs("ID") then
								response.Write("<option value='"&rs("ID")&"' selected>"&rs("Content")&"</option>")
								TowID=rs("ID")	
							else
								response.Write("<option value='"&rs("ID")&"'>"&rs("Content")&"</option>")	
							end if
						else
							response.Write("<option value='"&rs("ID")&"'>"&rs("Content")&"</option>")					
						end if
					end if
					rs.movenext
				wend
			end if
			CloseRs
		else
			TowID="Null"
			response.Write("<option value='Null'>暂无信息！</option>")
		end if
	End Sub
	
	Public Sub ThreeGread()
		if TowID<>"Null" and TowID<>"" and IsNumeric(TowID) then
			SetRs
			Sql="Select * from County where ParentID2="&TowID&" order by px desc,id asc"
			Rs.OPen Sql,conn,1,1
			if rs.eof and rs.bof then
				response.Write("<option value='Null'>暂无信息！</option>")
			else
				while not rs.eof
					if QY_Citys<>"" and IsNumeric(QY_Citys) and QY_Citys=rs("ID") then
						response.Write("<option value='"&rs("ID")&"' selected>"&rs("Content")&"</option>")					
					else
						if Request_Citys<>"" and IsNumeric(Request_Citys) then
							if  Cint(Request_Citys)=rs("ID") then
								response.Write("<option value='"&rs("ID")&"' selected>"&rs("Content")&"</option>")	
							else
								response.Write("<option value='"&rs("ID")&"'>"&rs("Content")&"</option>")
							end if
						else
							response.Write("<option value='"&rs("ID")&"'>"&rs("Content")&"</option>")					
						end if
					end if
					rs.movenext
				wend			
			end if
			CloseRs
		else
			response.Write("<option value='Null'>暂无信息！</option>")
		end if
	End Sub
End Class
%>