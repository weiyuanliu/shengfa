<title>�ϲ�������Ϣ</title>
<%
Server.ScriptTimeout=999
Dim Conn,Conn2

'On error resume next
Set Conn=Server.CreateObject("Adodb.Connection")
Conn.open "Driver={SQL Server};server=(local);uid=belojdbuser;pwd=1asf2^&*64afafAFAF;database=belojdb;"

Set Conn2=Server.CreateObject("Adodb.Connection")
Conn2.open "Driver={SQL Server};server=162.212.182.195;uid=belojdbuser;pwd=1asf2^&*64afafAFAF;database=belojdb;"

if err then
   err.clear
   Set Conn = Nothing
   Response.Write "ϵͳ�������ݿ����ӳ�������'ϵͳ����>>վ�㳣������',����/Include/Const.asp�ļ�!"
   Response.End
end if
'ͬ��������Ϣ
call getmessagelink()
Function getmessagelink()
	Dim rs,sql,StartDate,EndDate
	StartDate=now-3
	EndDate=now
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from NwebCn_Message where (addtime between '" & StartDate & "' and '" & Cdate(EndDate)+1 & "')"

	rs.open sql,conn,1,1
	if not rs.eof then
	  while not rs.eof
	    call getOtherMessage(rs("MesName"),rs("Content"),rs("linkman"),rs("Mobile"),rs("ViewFlag"),rs("SecretFlag"),rs("AddTime"),rs("ReplyContent"),rs("ReplyTime"))
	  rs.movenext
	  wend
	end if
	rs.close
	set rs=nothing
End Function
Function getOtherMessage(MesName,Content,linkman,Mobile,ViewFlag,SecretFlag,AddTime,ReplyContent,ReplyTime)
	Dim rs,sql,StartDate,EndDate
	StartDate=now-3
	EndDate=now
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from NwebCn_Message where MesName<>'"&MesName&"' and Mobile<>'"&Mobile&"' and (addtime between '" & StartDate & "' and '" & Cdate(EndDate)+1 & "')"
	rs.open sql,Conn2,1,1
	if not rs.eof then
	   response.Write rs("ID")&":"&rs("MesName")&"<br/>"
       end if
	rs.close
	set rs=nothing
End Function
%>