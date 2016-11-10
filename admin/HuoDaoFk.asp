<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七日科技企业网站管理系统（LISuo）　　　　　　　  ┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
' 　版权所有　qisehu.com
'
'　　程序制作　七日科技有限公司
'　　　　　　　Add:四川省成都市二环路西三段181号13楼20/21号
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%>
<% Option Explicit %>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>查看、修改、回复订单</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|81,")=0 then 
  response.write ("<script language=javascript> alert('你不具有该管理模块的操作权限，请返回！');history.back(-1);</script>")
end if
%>
<%
if Instr(session("AdminPurview"),"|94,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
Dim id,Action
Action=Trim(Request.QueryString("Action"))
id=trim(Request.QueryString("ID"))
if id="" or isnull(id) or not(IsNumeric(id)) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('数据出错，请返回！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
else
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select State,FuKuan from NwebCn_Order where id="&id
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('记录未找到，发货失败！');"&vbcrlf)
			response.Write("window.history.go(-1);")
		response.Write("</script>")
		response.End()
	else
		if rs("State")="货已发" then
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('货已发出，不能重复操作！');")
				response.Write("window.history.go(-1);")
			response.Write("</script>")
			response.End()
		else
			if Action="true" then
				rs("State")="货到后付款"
				rs("FuKuan")=true
			else
				rs("State")="不能发货"
				rs("FuKuan")=false
			end if
			rs.update()
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('操作成功！');"&vbcrlf)
				response.Write("window.location.href=document.referrer;")
			response.Write("</script>")
		end if
	end if
end if
%>
