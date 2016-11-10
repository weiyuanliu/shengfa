<%response.charset="gb2312"%>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
Dim ParentID
ParentID=Trim(Request("ParentID"))
if ParentID="" or isnull(ParentID) or Not(IsNumeric(ParentID)) then
	response.Write("$error$")
else
	response.Write(GetValue(ParentID))
end if 

function GetValue(ParentID)
	dim rs,sql,str
	str="$"
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from City where ParentID="&ParentID&" order by px asc,id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		str=str&"ÔÝÎÞÐÅÏ¢"
	else
		while not rs.eof
			str=str&rs("ID")&","&rs("Content")&"|"
			rs.movenext
		wend 
		str=mid(str,1,len(str)-1)
	end if
	str=str&"$"
	GetValue=str
end function
%>