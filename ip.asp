
<% response.charset="gb2312" %>
<%
'ASPץȡԶ��ҳ�湦���ࣨ�Զ��жϱ����ʽ��
Function GetHttpPage(murlip)
dim Http,ore,Matches
Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
Http.open "GET",murlip,False
Http.Send()
If Http.Readystate<>4 and Http.status<>200 then
Set Http=Nothing
Exit function
End if
Set ore = New RegExp
ore.Pattern = "<meta[^>]+charset=[""]?([\w\-]+)[^>]*>"
ore.Global = True
ore.IgnoreCase = True
Set Matches = ore.execute(Http.responseText)
If(Matches.count>0)Then
GetHTTPPage=bytesToBSTR(Http.responseBody,Matches(0).submatches(0))
Else  
'GetHTTPPage=Http.responseText  'û���ҵ�������ת������
GetHTTPPage=bytesToBSTR(Http.responseBody,"utf-8") 'û���ҵ�������ת��ΪGB2312
End if
Set Http=Nothing
End Function

Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText
objstream.Close
set objstream = nothing
End Function

Function GetKey(HTML,Start,Last)
dim filearray,filearray2
filearray=split(HTML,Start)
if ubound(filearray)>0 then
filearray2=split(filearray(1),Last)
GetKey=filearray2(0)
end if
End Function

dim bookip
bookip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
if bookip= "" Then bookip=Request.ServerVariables("REMOTE_ADDR")

'ip138 ip
dim murlip,StartGetip,iptoadd,ipto
murlip="http://www.ip.cn/index.php?ip="&bookip
StartGetip = getHTTPPage(murlip)
iptoadd=Getkey(StartGetip,"���ԣ�","</p>")
if iptoadd <> "" then
	ipto=iptoadd
end if
%>