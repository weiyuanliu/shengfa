<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
dim id,Zt
 id=Request.QueryString("ID")
 zt=Request.QueryString("zt")
  if zt="bl" then
    if id<>"" then  conn.execute "update NwebCn_Message set where id in ("&id&")"
  elseif zt="cg" then
   if id<>"" then  conn.execute "update NwebCn_Message set Flag=1 where id in ("&id&")"
  else
    if id<>"" then  conn.execute "update NwebCn_Message set Flag=0 where id in ("&id&")"
  end if
  
    response.redirect request.servervariables("http_referer")
%>