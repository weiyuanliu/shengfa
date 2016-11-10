<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->

<%
dim id,Zt,del
 id=Request.QueryString("ID")
 zt=Request.QueryString("zt")
 del=session("UserName")
  if zt="bl" then
    if id<>"" then  conn.execute "update NwebCn_Message set where id in ("&id&")"
  elseif zt="cg" then
   if id<>"" then  conn.execute "update NwebCn_Message set Flag=0,ViewFlag=1,deladmin='"&del&"' where id in ("&id&")"
  else
    if id<>"" then  conn.execute "update NwebCn_Message set Flag=1,ViewFlag=0,deladmin='"&del&"' where id in ("&id&")"
  end if
  
    response.redirect request.servervariables("http_referer")
%>