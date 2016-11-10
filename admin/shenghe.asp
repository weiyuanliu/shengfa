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
<TITLE>审核、修改、回复留言</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|92,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<%Call Set_TuiJian()%>
<%
	Sub Set_TuiJian()
		Dim ID,Rs,Sql
		ID=Trim(Request.QueryString("ID"))
		if ID="" or Isnull(ID) or Not(IsNumeric(ID)) then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('对不起，数据出错请返回！');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>"&vbCrlf)
			response.End()
			exit sub
		end if
		
		Set Rs=server.CreateObject("adodb.recordset")
		sql="select ViewFlag from NwebCn_Message where id="&id
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			rs.close()
			set rs=Nothing
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('记录未找到，操作失败！');"&vbcrlf)
				response.Write("window.history.go(-1);")
			response.Write("</script>"&vbcrlf)
			response.End()
			exit sub
		else
			if trim(Request.QueryString("Action"))="true" then 
				rs("ViewFlag")=1
			else
				rs("ViewFlag")=0
			end if
			rs.update()
		end if
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('操作成功！');")
			response.Write("window.location.href=document.referrer;")
		response.Write("</script>"&vbcrlf)
	End Sub
%>
</BODY>
</HTML>
