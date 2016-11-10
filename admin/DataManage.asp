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
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>数据库操作</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|310,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>数据库操作：系统数据备分，压缩，恢复，管理员登录日志</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9">
    <a href="DataManage.asp" onClick='changeAdminFlag("数据库操作")'>栏目首页</a><font color="#0000FF">&nbsp;|&nbsp;</font>网站数据库：<a href="DataManage.asp?Action=DataBackup&Result=Site" onClick='changeAdminFlag("网站数据库备份")'>备份</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataCompact&Result=Site" onClick='changeAdminFlag("网站压缩数据库")'>压缩</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataResume&Result=Site" onClick='changeAdminFlag("网站恢复数据库")'>恢复</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="DataManage.asp?Action=DataLog" onClick='changeAdminFlag("管理员登录日志")'>管理员登录日志</a><font color="#0000FF">&nbsp;|&nbsp;</font>流量数据库：<a href="DataManage.asp?Action=DataBackup&Result=Stat" onClick='changeAdminFlag("流量数据库备份")'>备份</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataCompact&Result=Stat" onClick='changeAdminFlag("流量压缩数据库")'>压缩</a>&nbsp;&nbsp;<a href="DataManage.asp?Action=DataResume&Result=Stat" onClick='changeAdminFlag("流量恢复数据库")'>恢复</a><font color="#0000FF">&nbsp;</font></td>    
  </tr>
</table>
<br>
<% call DataManage() %>
</body>
</html>
<%
sub DataManage()
  Dim Action
  Action=request.QueryString("Action")
  Select Case Action
    Case "DataBackup"
	  DataBackup
    Case "DataCompact"
	  DataCompact
    Case "DataResume"
	  DataResume
    Case "DataLog"
	  DataLog
    Case Else
      DataMain
  End Select
end sub  
%>

<%
function DataMain
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9'>")
  response.write ("操作说明：<br>　　·数据库操作步骤为[备份&nbsp;→&nbsp;压缩&nbsp;→&nbsp;恢复]<br>　　·操作前最好先[<font color='#330099'>备份</font>]数据库，正在使用中的数据库不能被压缩<BR>　　·恢复数据库时将会覆盖当前使用中的数据库<br>　　·管理员登录日志可做查看、删除")
  response.write ("</td></tr></table>")
end function

function DataBackup()
  dim From,Fso,Result
  From=request.QueryString("From")
  Result=request.QueryString("Result")
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9' align='center'>")
  response.write ("<table width='560' border='0' cellspacing='0' cellpadding='0'><tr><td height='16'></td></tr>")
  response.write ("<tr><td height='20'>说明：修改数据库备份保存路径和文件名，请进入[系统设置→站点常量设置→数据库备份路径]</td></tr>")
  if From="Confirm" then
    set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if Result="Site" then
	  Fso.CopyFile Server.MapPath(SiteDataPath),Server.MapPath(SiteDataBakPath)
      response.write ("<tr><td height='20'>成功：你已经成功备份数据到&nbsp;<a href='"&SiteDataBakPath&"' target='_blank'><font color='#330099'>"&SiteDataBakPath&"</font></a>&nbsp;，注意及时删除不用的备份！</td></tr>")
	else
	  Fso.CopyFile Server.MapPath(StatDataPath),Server.MapPath(StatDataBakPath)
      response.write ("<tr><td height='20'>成功：你已经成功备份数据到&nbsp;<a href='"&StatDataBakPath&"' target='_blank'><font color='#330099'>"&StatDataBakPath&"</font></a>&nbsp;，注意及时删除不用的备份！</td></tr>")
    end if
 	response.write ("<tr><td height='20'>版本：数据库的时间版本为&nbsp;"& now() &"</td></tr>")
    Set Fso=nothing
  end if	  
  response.write ("<form id='DataBackupForm' name='DataBackupForm' method='post' action='DataManage.asp?From=Confirm&Action=DataBackup&Result="&Result&"'>")
  if Result="Site" then
    response.write ("<tr><td height='30'>来源：<input name='fromPath' readonly type='text' size='60' value='"&SiteDataPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>目标：<input name='toPath' readonly type='text' size='60' value='"&SiteDataBakPath&"' class='textfield' /></td></tr>")
  else
    response.write ("<tr><td height='30'>来源：<input name='fromPath' readonly type='text' size='60' value='"&StatDataPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>目标：<input name='toPath' readonly type='text' size='60' value='"&StatDataBakPath&"' class='textfield' /></td></tr>")
  end if
  response.write ("<tr><td height='30'><input type='submit' value='确定备份' class='button' /></td></tr>")
  response.write ("</form>")  
  response.write ("<tr><td height='16'></td></tr></table>")
  response.write ("</td></tr></table>")
end function

function DataCompact()
  dim From,Fso,Engine,SDBPath,Result
  From=request.QueryString("From")
  Result=request.QueryString("Result")
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9' align='center'>")
  response.write ("<table width='560' border='0' cellspacing='0' cellpadding='0'><tr><td height='16'></td></tr>")
  response.write ("<tr><td height='20'>说明：压缩前最好先[<font color='#330099'>备份</font>]数据库，正在使用中的数据库不能被压缩</td></tr>")
  if From="Confirm" then
    if Result="Site" then
      SDBPath = server.mappath(SiteDataBakPath)
	else
      SDBPath = server.mappath(StatDataBakPath)
	end if
    set Fso=Server.CreateObject("Scripting.FileSystemObject")
	if Fso.FileExists(SDBPath) then
      Set Engine =Server.CreateObject("JRO.JetEngine")
	  if request("boolIs") = "97" then
	     Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath, _
		                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath & "_temp.mdb;" _
		                        & "Jet OLEDB:Engine Type=" & JET_3X
	  else 
	     Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath, _
		                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SDBPath & "_temp.mdb"
      end if
      Fso.CopyFile SDBPath & "_temp.mdb",SDBPath
      Fso.DeleteFile(SDBPath & "_temp.mdb")
      set Fso = nothing
      set Engine = nothing
      response.write ("<tr><td height='20'>成功：数据库&nbsp;<a href='"&SDBPath&"' target='_blank'><font color='#330099'>"&SiteDataBakPath&"</font></a>&nbsp;已经压缩成功！</td></tr>")
	  response.write ("<tr><td height='20'>版本：数据库的时间版本为&nbsp;"& now() &"</td></tr>")
    else
      response.write ("<tr><td height='20'>失败：数据库&nbsp;<a href='"&SDBPath&"' target='_blank'><font color='#330099'>"&SiteDataBakPath&"</font></a>&nbsp;压缩失败，请检查路径和数据库名是否存在！</td></tr>")
    end if
  end if
  response.write ("<form id='DataCompactForm' name='DataCompactForm' method='post' action='DataManage.asp?From=Confirm&Action=DataCompact&Result="&Result&"'>")
  if Result="Site" then
    response.write ("<tr><td height='30'>目标：<input name='toPath' readonly type='text' size='60' value='"&SiteDataBakPath&"' class='textfield'/></td></tr>")
  else
    response.write ("<tr><td height='30'>目标：<input name='toPath' readonly type='text' size='60' value='"&StatDataBakPath&"' class='textfield'/></td></tr>")
  end if
  response.write ("<tr><td height='30'><input type='submit' value='确定压缩' class='button' /></td></tr>")
  response.write ("</form>")  
  response.write ("<tr><td height='16'></td></tr></table>")
  response.write ("</td></tr></table>")
end function

function DataResume()
  dim From,Fso,SDPath,SDBPath,Result
  From=request.QueryString("From")
  Result=request.QueryString("Result")
  response.write ("<table width='100%' border='0' cellpadding='3' cellspacing='1' bgcolor='#6298E1'><tr><td height='24' nowrap  bgcolor='#EBF2F9' align='center'>")
  response.write ("<table width='560' border='0' cellspacing='0' cellpadding='0'><tr><td height='16'></td></tr>")
  response.write ("<tr><td height='20'>说明：修改备份、目标数据库的保存路径和文件名，请进入[系统设置→站点常量设置→数据库备份路径]</td></tr>")
  if From="Confirm" then
    if Result="Site" then
	  SDPath = server.mappath(SiteDataPath)
      SDBPath = server.mappath(SiteDataBakPath)
	else
	  SDPath = server.mappath(StatDataPath)
      SDBPath = server.mappath(StatDataBakPath)
	end if
    set Fso=Server.CreateObject("Scripting.FileSystemObject")
    if Fso.FileExists(SDBPath) then
      Fso.CopyFile SDBPath,SDPath
      Set Fso=nothing
      response.write ("<tr><td height='20'>成功：你已经成功恢复数据库&nbsp;<font color='#330099'>"&SDPath&"</font>&nbsp;注意及时删除不用的备份！</td></tr>")
	  response.write ("<tr><td height='20'>版本：数据库的时间版本为&nbsp;"& now() &"</td></tr>")
    else
      response.write ("<tr><td height='20'>失败：数据库&nbsp;<a href='"&SDBPath&"' target='_blank'><font color='#330099'>"&SDBPath&"</font></a>&nbsp;压缩失败，请检查路径和数据库名是否存在！</td></tr>")
    end if
  end if	    
  response.write ("<form id='DataResumeForm' name='DataResumeForm' method='post' action='DataManage.asp?From=Confirm&Action=DataResume&Result="&Result&"'>")
  if  Result="Site" then
    response.write ("<tr><td height='30'>来源：<input name='fromPath' readonly type='text' size='60' value='"&SiteDataBakPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>目标：<input name='toPath' readonly type='text' size='60' value='"&SiteDataPath&"' class='textfield' /></td></tr>")
  else
    response.write ("<tr><td height='30'>来源：<input name='fromPath' readonly type='text' size='60' value='"&StatDataBakPath&"' class='textfield'/></td></tr>")
    response.write ("<tr><td height='30'>目标：<input name='toPath' readonly type='text' size='60' value='"&StatDataPath&"' class='textfield' /></td></tr>")
  end if
  response.write ("<tr><td height='30'><input type='submit' value='确定恢复' class='button' /></td></tr>")
  response.write ("</form>")  
  response.write ("<tr><td height='16'></td></tr></table>")
  response.write ("</td></tr></table>")
end function

function DataLog()
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=LoginLog" method="post" name="formDel" >
    <tr>
      <td width="60" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
      <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>登录名</strong></font></td>
      <td width="70" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>用户名</strong></font></td>
      <td width="124" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong><font color="#FFFFFF">登录IP</font></strong></font></td>
      <td width="260" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">登录时浏览器</font></strong></td>
      <td width="124" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF"><strong>创建时间</strong></font></strong></td>
      <td nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">
	  </td>
    </tr>
	<% AdminLoginLog() %>
  </form>
</table>
<%
end function
%>
<%
'-----------------------------------------------------------
function AdminLoginLog()
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数
  dim page'页码
      page=clng(request("Page"))
  dim pagenc '每页显示的分页页码数量=pagenc*2+1
      pagenc=2
  dim pagenmax '每页显示的分页的最大页码
  dim pagenmin '每页显示的分页的最小页码
  dim datafrom'数据表名
      datafrom="NwebCn_AdminLog"
  dim datawhere'数据条件
      datawhere=""
  dim sqlid'本页需要用到的id
  dim Myself,PATH_INFO,QUERY_STRING'本页地址和参数
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" then
	    Myself = PATH_INFO & "?"
	  else
	  	if Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
          Myself = PATH_INFO & "?" & QUERY_STRING & "&"
		else
	      Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
		end if
	  end if
  dim taxis'排序的语句 asc, desc
      taxis="order by id desc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(ID) as idCount from ["& datafrom &"]" & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,0,1
  idCount=rs("idCount")
  '获取记录总数

  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select id from ["& datafrom &"] " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    rs.pagesize = pages '每页显示记录数
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("id")
	  else
	    sqlid=sqlid &","&rs("id")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  if(idcount>0 and sqlid<>"") then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select [ID],[AdminName],[UserName],[LoginIP],[LoginSoft],[LoginTime] from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("AdminName")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("UserName")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("LoginIP")&"</td>" & vbCrLf
	  if len(rs("LoginSoft"))>40 then
        Response.Write "<td nowrap title='浏览器：&#13;"&rs("LoginSoft")&"'>"&left(rs("LoginSoft"),40)&"...</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='浏览器：&#13;"&rs("LoginSoft")&"'>"&rs("LoginSoft")&"</td>" & vbCrLf
      end if 
      Response.Write "<td nowrap>"&rs("LoginTime")&"</td>" & vbCrLf
 	  Response.Write "<td nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='6' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='删除所选'  onClick='ConfirmDel(""您真的要删除这些管理员登录日志吗？"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write ("<tr><td height='50' align='center' colspan='7' nowrap  bgcolor='#EBF2F9'>暂无管理员登录日志</td></tr>")
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='7' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td>共计：<font color='#ff6600'>"&idcount&"</font>条记录&nbsp;页次：<font color='#ff6600'>"&page&"</font></strong>/"&pagec&"&nbsp;每页：<font color='#ff6600'>"&pages&"</font>条</td>" & vbCrLf
  Response.Write "<td align='right'>" & vbCrLf
  '设置分页页码开始===============================
  pagenmin=page-pagenc '计算页码开始值
  pagenmax=page+pagenc '计算页码结束值
  if(pagenmin<1) then pagenmin=1 '如果页码开始值小于1则=1
  if(page>1) then response.write ("<a href='"& myself &"Page=1'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>9</font></a>&nbsp;") '如果页码大于1则显示(第一页)
  if(pagenmin>1) then response.write ("<a href='"& myself &"Page="& page-(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>7</font></a>&nbsp;") '如果页码开始值大于1则显示(更前)
  if(pagenmax>pagec) then pagenmax=pagec '如果页码结束值大于总页数,则=总页数
  for i = pagenmin to pagenmax'循环输出页码
	if(i=page) then
	  response.write ("&nbsp;<font color='#ff6600'>"& i &"</font>&nbsp;")
	else
	  response.write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write ("&nbsp;<a href='"& myself &"Page="& page+(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>8</font></a>&nbsp;") '如果页码结束值小于总页数则显示(更后)
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>:</font></a>&nbsp;") '如果页码小于总页数则显示(最后页)	
  '设置分页页码结束===============================
  Response.Write "跳到：第&nbsp;<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('只能在跳转目标页框内输入整数！');this.value='"&Page&"';}"" style='HEIGHT: 18px;WIDTH: 40px;'  type='text' class='textfield' value='"&Page&"'>&nbsp;页" & vbCrLf
  Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick='GoPage("""&Myself&""")' value='GO'>" & vbCrLf
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf
  rs.close
  set rs=nothing
  Response.Write "</td>" & vbCrLf  
  Response.Write "</tr>" & vbCrLf
'-----------------------------------------------------------
'-----------------------------------------------------------
end function 
%>