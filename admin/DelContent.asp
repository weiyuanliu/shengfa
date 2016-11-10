<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->

<%
dim Result,Selectid
Result=request.QueryString("Result")
SelectID=request.Form("SelectID")
select case Result
  case "Administrators"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Admin where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Province"
  	if SelectID<>"" then  
		conn.execute "delete from Province where id in ("&SelectID&")"
		conn.execute "delete from City where ParentID in("&SelectID&")"
		conn.execute "delete from County where ParentID in("&SelectID&")"
		conn.execute "delete from Regional where QY_ShengFen in("&SelectID&")"
	end if
    response.redirect request.servervariables("http_referer")
  case "City"
  	if SelectID<>"" then  
		conn.execute "delete from City where id in ("&SelectID&")"
		conn.execute "delete from County where ParentID2 in("&SelectID&")"
		conn.execute "delete from Regional where QY_City in("&SelectID&")"
	end if
    response.redirect request.servervariables("http_referer")
  case "Regional"
  	if SelectID<>"" then  conn.execute "delete from Regional where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "County"
  	if SelectID<>"" then  
		conn.execute "delete from County where id in ("&SelectID&")"
		conn.execute "delete from Regional where QY_Citys in("&SelectID&")"
	end if
    response.redirect request.servervariables("http_referer")
  case "LoginLog"
    'if SelectID<>"" then  conn.execute "delete from NwebCn_AdminLog where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Members"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Members where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "About"
    if SelectID<>"" then  conn.execute "delete from NwebCn_About where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Products"
  	if selectId<>"" Then 
	DelPic "NwebCn_Products",Selectid 
	 conn.execute "delete from NwebCn_Products where id in ("&SelectID&")"
    end if
	response.redirect request.servervariables("http_referer")
  case "News"
  if selectId<>"" Then 
	  
	 conn.execute "delete from NwebCn_News where id in ("&SelectID&")"
	end if
  
    
    response.redirect request.servervariables("http_referer")
  case "Download"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Download where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Need"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Need where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "ADs"
    if SelectID<>"" then  conn.execute "delete from NwebCn_ADs where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Jobs"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Jobs where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Message"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Message where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "MsgData"
  	if SelectID<>"" then  conn.execute "delete from MsgData where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Order"
  
  
    if SelectID<>"" then  conn.execute "update   NwebCn_Order set fax=1 where id in ("&SelectID&")"
  
  
    response.redirect request.servervariables("http_referer")
   case "OrderD"
  
  
    if SelectID<>"" then  conn.execute "delete from   NwebCn_Order   where id in ("&SelectID&")"
  
  
    response.redirect request.servervariables("http_referer")
  
  case "Supply"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Supply where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Talents"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Talents where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "Navigation"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Navigation where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
  case "FriendSite"
    if SelectID<>"" then  conn.execute "delete from NwebCn_FriendSite where id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
	case "Others"
	if SelectID<>"" then  conn.execute "delete from NwebCn_Others where id in ("&SelectID&")"
	 response.redirect request.servervariables("http_referer")
  case "NoHackSql"
    if SelectID<>"" then  conn.execute "delete from NwebCn_NoHackSql where SqlIn_ID in ("&SelectID&")"
  
    response.redirect request.servervariables("http_referer")
  case "WAIBU_ADV"
    if SelectID<>"" then  conn.execute "delete from NwebCn_Ads_effect where Id in ("&SelectID&")"
    response.redirect request.servervariables("http_referer")
	
  case else
	
end select

Function DelFile(Files)
dim fs,file
Set fs = Server.CreateObject("Scripting.FileSystemObject")
if files<>"" then
File = Server.MapPath(Files)
on Error Resume Next
fs.DeleteFile File, True '强制删除只读文件
If Err.Number = 53 Then
Response.Write File & "文件不存在！"
Response.End
Elseif Err.Number = 70 Then
Response.Write File & "文件属性为锁定状态！"
Response.End
Elseif Err.Number <> 0 Then
Response.Write "未知错误，错误编码：" & Err.Number
Response.End
Else
Response.Write "成功删除文件！" & File
End If
end if
 
End Function

Function DelPic(Dates,Selectid)

dim sqls,rss
	set rss=server.CreateObject("Adodb.recordset")
	Sqls="select smallpic,bigpic from "&Dates&" where id in("&Selectid&") "
	 
	rss.open sqls,conn,1,3
	if rss.bof and rss.eof then
	else
	while not rss.eof
	if rss("smallpic")=rss("bigpic") then
	 DelFile(Rss("smallpic"))
	 else
	 
	 DelFile(Rss("bigpic"))
	end if
	rss.movenext
	wend
	end if
	rss.close
	set rss=nothing
End Function
%>