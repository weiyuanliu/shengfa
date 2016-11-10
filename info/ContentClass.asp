<%
Class TextClass
	Dim Rs,Sql,ID,ZiDuan,TableName
    
    Private Sub SetRs()
    	Set Rs=Server.Createobject("Adodb.Recordset")
    End Sub
	
	Private Sub CloseRs()
		Rs.Close()
		set Rs=Nothing
	End Sub
	
	Public Sub Set_ID(Values)
		ID=Values
	End Sub
	
	Public Sub Set_ZiDuan(Str)
		ZiDuan=Str
	End Sub
	
	Public Sub Set_TableName(Values)
		TableName=Values
	End sub
		
	Private Function Get_Value() '获取指定数据库表中指定ID字段的值
		SetRs
		Sql="Select "&ZiDuan&" from "&TableName&" Where ViewFlag=1 and ID="&ID
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			Get_Value="内容添加中，请稍后……"
		else
			Get_Value=rs(ZiDuan)
		end if
		CloseRs
	End Function
	
	Public Sub Print_Text()
		Response.Write(Get_Value)
	End Sub
End Class
%>