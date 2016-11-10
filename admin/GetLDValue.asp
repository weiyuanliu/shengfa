<%response.charset="gb2312"%>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
	Dim Action,ParentID
	Action=Trim(Request("Action"))
	ParentID=Trim(Request("ParentID"))
	if Action="" or isnull(Action) or ParentID="" or isnull(ParentID) or Not(IsNumeric(ParentID)) then
		response.Write("$Error$")
	else
		response.Write(Get_Value(Action,ParentID))
	end if
	
	Function Get_Value(Action,ParentID)
		Dim Rs,Sql,Str,NextGrad
		Set Rs=server.CreateObject("adodb.recordset")
		Str="$"
		if Action="Two" then
			Sql="Select * from City where ParentID="&ParentID&" order by px desc,id asc"
			Rs.Open Sql,Conn,1,1
			if Rs.eof and rs.bof then
				Str=Str&"Null,暂无信息！||Null,暂无信息！|"
			else
				NextGrad=rs("ID")
				While not Rs.eof 
					Str=Str&rs("ID")&","&rs("Content")&"|"
					Rs.movenext
				Wend
				Str=Mid(Str,1,Len(Str)-1)
				rs.close()
				sql="select * from County where ParentID2="&NextGrad&" order by px desc,id asc"
				rs.open sql,conn,1,1
				if rs.eof and rs.bof then
					Str=Str&"||"&"Null,暂无信息！|"
				else
					Str=Str&"||"
					while not rs.eof
						Str=Str&rs("ID")&","&rs("Content")&"|"
						rs.movenext
					wend
				end if
			end if
			rs.close()
			Set rs=Nothing
		elseif Action="Three" then
			Sql="Select * from County where ParentID2="&ParentID&" order by px desc,id asc"
			rs.open sql,conn,1,1
			if rs.eof and rs.bof then
				Str=Str&rs("ID")&","&rs("Content")&"|"
			else
				while not rs.eof
					str=str&rs("ID")&","&rs("Content")&"|"
					rs.movenext
				wend
			end if
			rs.close()
			set rs=Nothing
		else
			Str=Str&"Error"
		end if
		Str=Str&"$"
		Get_Value=Str
	End Function
%>