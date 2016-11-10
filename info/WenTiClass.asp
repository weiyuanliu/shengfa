<%
Class WenTiClass
	Dim Rs,Sql,ID,TableName
    
    Private Sub SetRs()
    	set rs=server.createobject("adodb.recordset")
    End Sub
    Private Sub CloseRs()
    	rs.close()
        set rs=Nothing
    End sub
    Public Sub Set_ID(Values)
    	ID=Values
    End Sub
    Public Sub Set_TableName(Values)
    	TableName=Values
    End Sub
    Public Sub PrintText()
    	SetRs
    	Sql="Select Content from "&TableName&" where id="&ID&" and ViewFlag"
        Rs.open Sql,conn,1,1
        if Rs.eof and Rs.bof then
        	response.write("内容添加中，请稍后……")
        Else
        	response.write(ReStrReplace(rs("Content")))
        End IF
        CloseRs
    End Sub
End Class
%>