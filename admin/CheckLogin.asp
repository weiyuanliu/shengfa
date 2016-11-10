<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七色狐企业网站管理系统（qisehu）　　　　　　　  ┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
' 　版权所有　qisehu.com
'
'　　程序制作　七色狐网络有限公司
'　　　　　　　Add:四川省成都市二环路西三段181号13楼20/21号
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - qisehu.com-STUDIO" />
<META NAME="Author" CONTENT="顺意网络有限公司" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>检查管理登录</TITLE>
</HEAD>
<!--#include file="../Include/Const.asp"-->
<!--#include file="../Include/ConnSiteData.asp"-->
<!--#include file="../Include/Md5.asp"-->
<BODY>
<%
'if request("t") <> "Login" then
'  response.Redirect("Login.asp")
'  response.End()
'end if

dim Id,LoginName,LoginPassword,AdminName,Password,AdminPurview,Working,UserName,GroupID,rs,sql
LoginName=trim(request.form("LoginName"))
LoginPassword=Md5(request.form("LoginPassword"))
set rs = server.createobject("adodb.recordset")
sql="select * from NwebCn_Admin where AdminName='"&LoginName&"'"
rs.open sql,conn,1,3

if rs.eof then
   response.write "<script language=javascript> alert('管理员名称不正确，请重新输入!');location.replace('Login.asp');</script>"
   response.end
else
   Id=rs("Id")
   AdminName=rs("AdminName")
   Password=rs("Password")
   GroupID=rs("GroupID")
   AdminPurview=rs("AdminPurview")
   Working=rs("Working")
   UserName=rs("UserName")
end if

if LoginPassword<>Password then
   response.write "<script language=javascript> alert('管理员密码不正确，请重新输入!!');location.replace('Login.asp');</script>"
   response.end
end if 

if session("VerifyCode")<>request("VerifyCode") then
   response.write "<script language=javascript> alert('您输入验证码错误，请返回重新登录！');location.replace('Login.asp');</script>"
   response.end
end if

if not Working then
   response.write "<script language=javascript> alert('不能登录，此管理员帐号已被锁定。');location.replace('Login.asp');</script>"
   response.end
end if 
 
if LoginName=AdminName and LoginPassword=Password then
   rs("LastLoginTime")=now()
   rs("LastLoginIP")=Request.ServerVariables("Remote_Addr")
   rs.update
   rs.close
   set rs=nothing 
   session("AdminName")=AdminName
   session("AdminId")=Id
   session("GroupID")=GroupID
   session("UserName")=UserName
   session("AdminPurview")=AdminPurview
   session("LoginSystem")="Succeed"

   response.Cookies("AdminName")=AdminName
   response.Cookies("AdminId")=Id
   response.Cookies("UserName")=UserName
   response.Cookies("AdminPurview")=AdminPurview
   response.Cookies("LoginSystem")="Succeed"
   session.timeout=30
   '==================================
    dim LoginIP,LoginTime,LoginSoft
   LoginIP=Request.ServerVariables("Remote_Addr")
   LoginSoft=Request.ServerVariables("Http_USER_AGENT")
   LoginTime=now()
   '====================================
   set rs = server.createobject("adodb.recordset")
   sql="select * from NwebCn_AdminLog"
   rs.open sql,conn,1,3
   rs.addnew
   rs("AdminName")=AdminName
   rs("UserName")=UserName
   rs("LoginIP")=LoginIP
   rs("LoginSoft")=LoginSoft
   rs("LoginTime")=LoginTime
   rs.update
   rs.close
   set rs=nothing 
   '========================================
   response.redirect "main.asp"
   response.end
end if
%>
</BODY>
</HTML>