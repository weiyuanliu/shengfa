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
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
</HEAD>
<!--#include file="CheckAdmin.asp"-->
<BODY>
<div align="center"><table width="720" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td width="50%" height="24" bgcolor="#EBF2F9">qisehu当前使用版本：LS2007 Build 0518 </td>
    <td width="50%" height="24" bgcolor="#EBF2F9">当前官方版本：LS2007 Build 0518 </td>
  </tr>
</table>
<br>
<table width="720" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" colspan="2"><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>服务器信息</strong></font></td>
    </tr>
  <tr>
    <td width="50%" height="24" bgcolor="#EBF2F9">服务器操作系统：<%=Request.ServerVariables("OS")%></td>
    <td width="50%" height="24" bgcolor="#EBF2F9">网站信息服务软件和版本：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr>
    <td width="50%" height="24" bgcolor="#EBF2F9">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
    <td width="50%" height="24" bgcolor="#EBF2F9">脚本超时时间：<%=Server.ScriptTimeout%>秒</td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">CDONTS组件支持：<%
	  On Error Resume Next
	  Server.CreateObject("CDONTS.NewMail")
	  if err=0 then 
		 response.write("<font color=red>√</font>")
	  else
         response.write("<font color=red>×</font>")
	  end if
	  err=0
    %></td>
    <td height="24" bgcolor="#EBF2F9">Jmail邮箱组件支持：<%
	  If Not IsObjInstalled(theInstalledObjects(13)) Then
         response.write("<font color=red>×</font>") 
      else
         response.write("<font color=red>√</font>") 
      end if
    %></td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">返回服务器处理请求的端口：<%=Request.ServerVariables("SERVER_PORT")%></td>
    <td height="24" bgcolor="#EBF2F9">协议的名称和版本：<%=Request.ServerVariables("SERVER_PROTOCOL")%></td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">服务器 CPU 数量：<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%></td>
    <td height="24" bgcolor="#EBF2F9">FSO文本文件读写：<%
	On Error Resume Next
	Server.CreateObject("Scripting.FileSystemObject")
	if err=0 then 
	   response.write("<font color=red>√</font>，支持")
	else
       response.write("<font color=red>×</font>，不支持")
	end if 
	err=0
    %></td>
  </tr>
  <tr>
    <td height="24" bgcolor="#EBF2F9">客户端操作系统：<%
      dim thesoft,vOS
      thesoft=Request.ServerVariables("HTTP_USER_AGENT")
      if instr(thesoft,"Windows NT 5.0") then
	     vOS="Windows 2000"
      elseif instr(thesoft,"Windows NT 5.2") then
	     vOs="Windows 2003"
      elseif instr(thesoft,"Windows NT 5.1") then
         vOs="Windows XP"
      elseif instr(thesoft,"Windows NT") then
       	 vOs="Windows NT"
      elseif instr(thesoft,"Windows 9") then
	     vOs="Windows 9x"
      elseif instr(thesoft,"unix") or instr(thesoft,"linux") or instr(thesoft,"SunOS") or instr(thesoft,"BSD") then
	     vOs="类Unix"
      elseif instr(thesoft,"Mac") then
	     vOs="Mac"
      else
     	vOs="Other"
      end if
      response.Write(vOs)
    %></td>
    <td height="24" bgcolor="#EBF2F9">站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
  </tr>
  <tr>
    <td width="50%" height="24" bgcolor="#EBF2F9">域名IP：http://<%=Request.ServerVariables("SERVER_NAME")%>&nbsp;/&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
    <td width="50%" height="24" bgcolor="#EBF2F9">虚拟路径：<%=Request.ServerVariables("SCRIPT_NAME")%></td>
  </tr>
  <tr>
    <td height="24" colspan="2" bgcolor="#D7E4F7">客户端浏览器要求： IE5.5或以上，并关闭所有弹窗的阻拦程序；服务器建议采用：Windows 2000或Windows 2003 Server。</td>
    </tr>
</table>
<br>
<table width="720" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" colspan="4"><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>系统开发，版权所有，授权使用</strong></font></td>
    </tr>
  <tr>
    <td width="10%" height="24" bgcolor="#D7E4F7">授权使用：</td>
    <td width="40%" height="24" bgcolor="#D7E4F7">七日科技</td>
    <td width="10%" height="24" bgcolor="#D7E4F7">版权所有</td>
    <td width="40%" height="24" bgcolor="#D7E4F7"> 2004-2008 CopyRight <a href="http://www.qisehu.com" target="_blank">qisehu</a> Co.,LTD</td>
  </tr>
</table>
<%
if request.QueryString("Action")="save" then

 call saveedit()
end if
sub saveedit()
dim rspur,sqlpur,leftpur
   set Rspur=server.CreateObject("Adodb.recordset")
   sqlpur="select * from Purview where id=1"
   rspur.open sqlpur,conn,1,3
   rspur("qxsz")=Request.QueryString("qxsz")
   rspur("leftPurview")=Request.Form("leftPur11") & Request.Form("leftPur12") & Request.Form("leftPur21") & Request.Form("leftPur22") & Request.Form("leftPur23") &Request.Form("leftPur31") & Request.Form("leftPur32") & Request.Form("leftPur33") &Request.Form("leftPur41") & Request.Form("leftPur42") & Request.Form("leftPur43") &Request.Form("leftPur51") & Request.Form("leftPur52") & Request.Form("leftPur53") &Request.Form("leftPur61") & Request.Form("leftPur62") &Request.Form("leftPur71") & Request.Form("leftPur72") & Request.Form("leftPur73") &Request.Form("leftPur81") & Request.Form("leftPur82") &Request.Form("leftPur91") & Request.Form("leftPur92")& Request.Form("leftPur93")
	rsPur.Update
	rspur.close
	set rspur=Nothing
			
end sub
dim rspur,sqlpur,leftpur
   set Rspur=server.CreateObject("Adodb.recordset")
   sqlpur="select * from Purview"
   rspur.open sqlpur,conn,1,3
   if rspur.bof and rspur.eof then 
   Response.Write("记录不存在")
   else
   
  
   if rspur("qxsz")=0 then 
   leftpur=rspur("leftPurview")
else
   leftpur=rspur("leftPurview")
%>
<br>
<table width="720" border="0" cellpadding="3" cellspacing="1" bgcolor="#639AE7">
<form name="sysform" method="post" action="syscome.asp?qxsz=1&Action=save">
  <tr>
    <td height="24" colspan="4"><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>系统设置</strong></font></td>
    </tr>
  <tr>
    <td width="69" height="20" align="right" bgcolor="#EFF3FF">操作权限：</td>
        <td width="636" nowrap bgcolor="#EFF3FF">
		  <input name="leftpur11" type="checkbox" value="|11," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|11,")>0 then response.write ("checked")%>>
		  企业信息
          <input name="leftpur12" type="checkbox" value="|12," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|12,")>0 then response.write ("checked")%>>
          新闻资讯
		  <input name="leftpur21" type="checkbox" value="|21," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|21,")>0 then response.write ("checked")%>>
		  新闻类别
		  <input name="leftpur22" type="checkbox" value="|22," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|22,")>0 then response.write ("checked")%>> 
		  产品展示
          <input name="leftpur23" type="checkbox" value="|23," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|23,")>0 then response.write ("checked")%>> 
          产品类别
		  <input name="leftpur31" type="checkbox" value="|31," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|31,")>0 then response.write ("checked")%>>
		  &nbsp;需求信息
		  <input name="leftpur32" type="checkbox" value="|32," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|32,")>0 then response.write ("checked")%>>
		  需求类别
		  <input name="leftpur33" type="checkbox" value="|33," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|33,")>0 then response.write ("checked")%>>
		  &nbsp;下载信息</td>
    </tr>
      <tr >
        <td height="20" align="right" bgcolor="#EFF3FF">&nbsp;</td>
        <td bgcolor="#EFF3FF">
		  <input name="leftpur41" type="checkbox" value="|41," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|41,")>0 then response.write ("checked")%>>
		  下载类别
		  <input name="leftpur42" type="checkbox" value="|42," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|42,")>0 then response.write ("checked")%>>
		  人才招聘
          <input name="leftpur43" type="checkbox" value="|43," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|43,")>0 then response.write ("checked")%>>
          其他信息
		  <input name="leftpur51" type="checkbox" value="|51," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|51,")>0 then response.write ("checked")%>>
		  &nbsp;其他类别
		  <input name="leftpur52" type="checkbox" value="|52," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|52,")>0 then response.write ("checked")%>>
		  弹出广告
          <input name="leftpur53" type="checkbox" value="|53," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|53,")>0 then response.write ("checked")%>>
          访客反馈
          <input name="leftpur61" type="checkbox" value="|61," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|61,")>0 then response.write ("checked")%>>
          留言信息
	      <input name="leftpur62" type="checkbox" value="|62," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|62,")>0 then response.write ("checked")%>>
	      &nbsp;定单信息</td>
      </tr>
      <tr >
        <td height="9" align="right" bgcolor="#EFF3FF">&nbsp;</td>
        <td bgcolor="#EFF3FF">
		  <input name="leftpur71" type="checkbox" value="|71," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|71,")>0 then response.write ("checked")%>>
		  供应信息
		  <input name="leftpur72" type="checkbox" value="|72," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|72,")>0 then response.write ("checked")%>>
		  人才信息
          <input name="leftpur82" type="checkbox" value="|82," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|82,")>0 then response.write ("checked")%>>
          用户管理
          <input name="leftpur73" type="checkbox" value="|73," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|73,")>0 then response.write ("checked")%>>
          会员管理
		  <input name="leftpur81" type="checkbox" value="|81," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|81,")>0 then response.write ("checked")%>>
		  会员组别
		  <input name="leftpur91" type="checkbox" value="|91," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|91,")>0 then response.write ("checked")%>>
		  网站管理员
		  <input name="leftpur92" type="checkbox" value="|92," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|92,")>0 then response.write ("checked")%>>
		  友情连接 
	      <label>
        <input name="leftpur93" type="checkbox" value="|93," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(leftpur,"|93,")>0 then response.write ("checked")%>>
&nbsp;分页条数 </label></td>
      </tr>
      <tr >
        <td height="10" align="right" bgcolor="#EFF3FF">&nbsp;</td>
        <td bgcolor="#EFF3FF"><input type="submit" name="Submit" value="保存"></td>
      </tr>
    </form>
</table>
<%
end if
 end if
 %>
</div>
</BODY>
</HTML>