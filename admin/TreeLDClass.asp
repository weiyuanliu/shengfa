<%
Class LdClass
	Dim Rs,Sql,FirstID,TowID,ID,QY_ShengFen,QY_City,QY_Citys
    Private Sub SetRs()
		Set Rs=Server.CreateObject("adodb.Recordset")
	End Sub
	
	Private Sub CloseRs()
		Rs.Close()
		Set Rs=Nothing
	End Sub
	
	Public Sub Set_QY_ShengFen(Values)
		QY_ShengFen=values
	End Sub
	
	Public Sub Set_QY_City(Values)
		QY_City=Values
	End Sub
	
	Public Sub Set_QY_Citys(Values)
		QY_Citys=Values
	End Sub
	
	Public Sub Set_ID(Values)
		ID=Values
	End Sub
	
	Public Sub FirstGread()
		SetRs
		Sql="Select * From Province order by Px desc,ID asc"
		Rs.open Sql,Conn,1,1
		if Rs.eof And Rs.bof then
			response.Write("<option value='Null'>������Ϣ��</option>")
		else
			FirstID=Rs("ID")
			While not rs.eof
				if QY_ShengFen<>"" and IsNumeric(QY_ShengFen) and QY_ShengFen=rs("ID") then
					response.Write("<option value='"&Rs("ID")&"' selected>"&rs("Content")&"</option>")
					FirstID=Rs("ID")
				else
					response.Write("<option value='"&Rs("ID")&"'>"&rs("Content")&"</option>")
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
				response.Write("<option value='Null'>������Ϣ��</option>")
			else
				TowID=rs("ID")
				while not rs.eof
					if QY_City<>"" and IsNumeric(QY_City) and QY_City=rs("ID") then
						response.Write("<option value='"&rs("ID")&"' selected>"&rs("Content")&"</option>")	
						TowID=rs("ID")			
					else
						response.Write("<option value='"&rs("ID")&"'>"&rs("Content")&"</option>")					
					end if
					rs.movenext
				wend
			end if
			CloseRs
		else
			TowID="Null"
			response.Write("<option value='Null'>������Ϣ��</option>")
		end if
	End Sub
	
	Public Sub ThreeGread()
		if TowID<>"Null" and TowID<>"" and IsNumeric(TowID) then
			SetRs
			Sql="Select * from County where ParentID2="&TowID&" order by px desc,id asc"
			Rs.OPen Sql,conn,1,1
			if rs.eof and rs.bof then
				response.Write("<option value='Null'>������Ϣ��</option>")
			else
				while not rs.eof
					if QY_Citys<>"" and IsNumeric(QY_Citys) and QY_Citys=rs("ID") then
						response.Write("<option value='"&rs("ID")&"' selected>"&rs("Content")&"</option>")					
					else
						response.Write("<option value='"&rs("ID")&"'>"&rs("Content")&"</option>")					
					end if
					rs.movenext
				wend			
			end if
			CloseRs
		else
			response.Write("<option value='Null'>������Ϣ��</option>")
		end if
	End Sub
End Class
%>