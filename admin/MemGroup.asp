<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'┌┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┐
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'┊　　　　　　　七日科技企业网站管理系统（ＮＷＥＢ）　　　　　　　┊
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'　　版权所有　lisuo.com
'
'　　程序制作　七日科技工作室
'　　　　　　　Add:四川省彭州市西大街228号/611930
'　　　　　　　Tel:028-68067902  Fax:83708850
'　　　　　　　E-m:duolaimi-123@163.com
'　　　　　　　Q Q:59309100
'
'　　相关网址　[产品介绍]http://www.qisehu.com
'　　　　　　　[支持论坛]http://www.qisehu.com/bbs
'
'　　演示网址　http://www.qisehu.com
'┊　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　┊
'└┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┘
%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>管理员组别</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312" />
<META NAME="copyright" CONTENT="Copyright 2004-2008 - lisuo.com-STUDIO" />
<META NAME="Author" CONTENT="成都七日科技有限公司,www.qisehu.com" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|105,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>网站会员管理：会员组别设置，添加，修改会员信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="MemGroup.asp?Result=Add" onClick='changeAdminFlag("添加管理员组别")'>添加管理员组别</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="AdminList.asp" onClick='changeAdminFlag("查看所有管理员")'>查看所有管理员</a></td>    
  </tr>
</table>
<br>
<% 
dim Result
Result=request.QueryString("Result")
dim ID,GroupID,GroupName,GroupLevel,Explain,AddTime,RanNum
ID=request.QueryString("ID")
randomize timer
RanNum=Int((8999)*Rnd +1009)
if Result<>"" then
  call MemGroupEdit() 
end if
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=MemGroup" method="post" name="formDel" >
    <tr>
      <td width="30" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
      <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>管理员组别号</strong></font></td>
      <td width="68" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>组别名称</strong></font></td>
      <td nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">说明</font></strong></td>
      <td width="118" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF"><strong>创建时间</strong></font></strong></td>
      <td width="76" colspan="2" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">      </td>
    </tr>
	<% MemGroupList() %>
  </form>
</table>
</BODY>
</HTML>
<%
sub MemGroupEdit()
  dim Action,rs,sql
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑会员组别信息
    set rs = server.createobject("adodb.recordset")
    if Result="Add" then '创建会员组别
	  sql="select * from NwebCn_MemGroup"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language=javascript> alert('会员组别名称必填，且字符数为6-16字符，3-8个汉字！');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupID")=Request.Form("GroupID")
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_MemGroup where ID="&ID
      rs.open sql,conn,1,3
      if len(trim(Request.Form("GroupName")))<3 or len(trim(Request.Form("GroupName")))>16  then
        response.write "<script language=javascript> alert('管理员组别名称必填，且字符数为6-16字符，3-8个汉字！');history.back(-1);</script>"
        response.end
      end if
	  rs("GroupName")=trim(Request.Form("GroupName"))
	  rs("Explain")=trim(Request.Form("Explain"))
      conn.execute("Update NwebCn_Members set GroupName='"&trim(Request.Form("GroupName"))&"' where GroupID='"&trim(Request.Form("GroupID"))&"'")
	end if
	rs.update
	rs.close
    set rs=nothing 
    response.write "<script language=javascript> alert('成功编辑管理员组别！');changeAdminFlag('管理员员组别');location.replace('MemGroup.asp');</script>"
  else '提取管理员信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_MemGroup where ID="& ID
      rs.open sql,conn,1,1
	  if rs.RecordCount=0 then
        response.write "<script language=javascript> alert('数据库中无此记录，请确定返回！');history.back(-1)</script>"
        response.end
	  end if
	  ID=rs("ID")
      GroupID=rs("GroupID")
	  GroupName=rs("GroupName")
	  Explain=rs("Explain")
	  rs.close
      set rs=nothing 
	end if
  end if
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editMemGroup" method="post" action="MemGroup.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onSubmit="return CheckMemGroup()">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">

      <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">I D：</td>
        <td><input name="ID" type="text" class="textfield" id="ID" style="WIDTH: 100;" value="<%if ID="" then response.write ("自动") else response.write (ID) end if%>" maxlength="6" readonly></td>
      </tr>
      <tr>
        <td height="20" align="right">管理员组号：</td>
        <td><input name="GroupID" type="text" class="textfield" id="GroupID" style="WIDTH: 100;" value="<%=GroupID%>" maxlength="4" >&nbsp;*组别号必须为数字，而且唯一</td>
      </tr>
      <tr>
        <td height="20" align="right">组别名称：</td>
        <td><input name="GroupName" type="text" class="textfield" id="GroupName" style="WIDTH: 100;" value="<%=GroupName%>">&nbsp;*管理员组别名称必填，且字符数为6-16字符，3-8个汉字！</td>
      </tr>
      <tr>
        <td height="20" align="right" valign="top">备注说明：</td>
        <td><textarea name="Explain" cols="88" rows="3" class="textfield" id="Explain" style="WIDTH: 580;" ><%=Explain%></textarea></td>
      </tr>

      <tr>
        <td height="30" align="right">&nbsp;</td>
        <td valign="bottom"><input name="submitSaveEdit" type="submit" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 60;" ></td>
      </tr>
      <tr>
        <td height="20" align="right">&nbsp;</td>
        <td valign="bottom">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  </form>
</table>
<br>
<%
end sub
'-----------------------------------------------------------
function MemGroupList()
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
      datafrom="NwebCn_MemGroup"
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
      taxis="order by id asc"
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
    sql="select * from ["& datafrom &"] where id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("GroupID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("GroupName")&"</td>" & vbCrLf
	  if len(rs("Explain"))>24 then
        Response.Write "<td nowrap title='说明：&#13;"&rs("Explain")&"'>"&left(rs("Explain"),22)&"...</td>" & vbCrLf
      else
        Response.Write "<td nowrap title='说明：&#13;"&rs("Explain")&"'>"&rs("Explain")&"</td>" & vbCrLf
      end if 
      Response.Write "<td nowrap>"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td width='40' nowrap><a href='MemGroup.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""修改管理员组别"")'><font color='#330099'>修改</font></a></td>" & vbCrLf
      if rs("ID")=1 then
	    Response.Write "<td width='22' nowrap></td>" & vbCrLf
      else
 	    Response.Write "<td width='22' nowrap><input name='selectID' type='checkbox' value='"&rs("GroupID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      end if
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='5' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td nowrap colspan='2'  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='删除所选'  onClick='ConfirmDel(""您真的要删除这些管理员组别吗？"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write ("<tr><td height='50' align='center' colspan='8' nowrap  bgcolor='#EBF2F9'>暂无会员组别</td></tr>")
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