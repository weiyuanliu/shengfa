<%
Class LxwmClass
	Dim rs,sql,Id,TableName
	
	Private Sub SetRs()
		Set rs=server.createobject("adodb.recordset")
	End Sub
	
	Private Sub CloseRs()
		rs.close()
		Set rs=Nothing
	End Sub
	
	Public Sub Set_TableName(Values)
		TableName=Values
	End Sub
	
	Public Sub Set_ID(Values)
		Id=Values
	End Sub
	
	Public Sub PrintText()
		SetRs
		Sql="Select Content From "&TableName&" where id="&Id&" and ViewFlag=1"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write("对不起，暂无信息！")
		else
			if rs("Content") <> "" then
			response.write(Replace(rs("Content"),"../",""))
			end if
		end if
		CloseRs
	End Sub
	Public Sub Printabout()
		SetRs
		Sql="Select aboutName From "&TableName&" where id="&Id&" and ViewFlag=1"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write("对不起，暂无信息！")
		else
			if rs("aboutName") <> "" then
			response.write(rs("aboutName"))
			end if
		end if
		CloseRs
	End Sub
End Class
%>