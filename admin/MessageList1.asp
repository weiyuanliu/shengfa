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
<TITLE>留言信息列表</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<style type="text/css">
<!--
.STYLE1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="select_date.asp"-->
<%
if Instr(session("AdminPurview"),"|91,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<%
dim Result,StartDate,EndDate,Keyword
Result=request.QueryString("Result")
StartDate=request.QueryString("StartDate")
EndDate=request.QueryString("EndDate")
Keyword=request.QueryString("Keyword")
function PlaceFlag()
  if Result="Search" then
    Response.Write "留言：列表&nbsp;->&nbsp;检索&nbsp;->&nbsp;留言时间[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，关键字[<font color='red'>"&Keyword&"</font>]"
  else
    Response.Write "留言：列表&nbsp;->&nbsp全部"
  end if
end function  
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>留言信息：审核，修改，回复留言信息相关的内容</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="Search.asp?Result=Message">
          <td nowrap> 留言检索：从
<%
	if Result="Search" then
		Response.Write "<input name=""start_date"" type=""text"" class=""textfield"" value="&StartDate&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date"" />到<input name=""end_date"" type=""text"" class=""textfield"" value="&EndDate&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date2"" />"
	else
		Response.Write "<input name=""start_date"" type=""text"" class=""textfield"" value="&dateadd("yyyy",-1,date())&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date"" />到<input name=""end_date"" type=""text"" class=""textfield"" value="&date()&" size=""10"" onfocus=""javascript:ShowCalendar(this.id)"" id=""select_date2"" />"
	end if
%>
          <!--<script language=javascript> 
          var myDate=new dateSelector(); 
          myDate.year--; 
		  myDate.date; 
          myDate.inputName='start_date';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>
          &nbsp;到
          <script language=javascript> 
          myDate.year++; 
          myDate.inputName='end_date';  //注意这里设置输入框的name，同一页中的日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>-->
          &nbsp;&nbsp;关键字：<input name="Keyword" type="text" class="textfield" value="<%=Keyword%>" size="18">
          <input name="submitSearch" type="submit" class="button" value="检索">
          </td>
        </form>
        <td align="right" nowrap>查看：<a href="MessageList.asp" onClick='changeAdminFlag("留言信息列表")'>留言列表</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SetSite.asp#Message" target="mainFrame" onClick='changeAdminFlag("网站信息设置")'>设置自动审核</a></td>
      </tr>
    </table>      </td>    
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"><%PlaceFlag()%></td>
  </tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=Message" method="post" name="formDel" >
  <tr>
    <td width="30" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
    <td width="88" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>留言地区</strong></font></td>
    <td width="193" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>留言标题</strong></font></td>
    <td width="194" nowrap bgcolor="#8DB5E9"><span class="STYLE1">地区</span></td>
    <td width="49" align="center" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF"><strong>标记</strong></font></strong></td>
    <td width="100" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">留言时间</font></strong></td>
    <td width="52" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审核</strong></font></td>
    <td width="118" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">回复时间</font></strong></td>
    <td colspan="2" width="76" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSelect" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">	</td>
  </tr>
  <% MessageList() %>
  </form>
</table>
</BODY>
</HTML>
<%
'-----------------------------------------------------------
function MessageList()
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数
  dim page'页码
      page=clng(request("Page"))
  dim pagenc'每页显示的分页页码数量=pagenc*2+1
      pagenc=2
  dim pagenmax'每页显示的分页的最大页码
  dim pagenmin'每页显示的分页的最小页码
  dim datafrom'数据表名
      datafrom="NwebCn_Message"
  dim datawhere'数据条件
      if Result="Search" then
	     datawhere="where ( MesName like '%" & Keyword &_
		           "%') and ( addtime between '" & StartDate & "' and '" & EndDate & "')"
	  else
        datawhere=""
	  end if
  dim sqlid'本页需要用到的id
  dim Myself,PATH_INFO,QUERY_STRING'本页地址和参数
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'排序的语句 asc,desc
      taxis="order by id desc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(ID) as idCount from ["& datafrom &"] " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,3
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
    rs.open sql,conn,1,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&Guest(rs("MemID"),rs("Linkman"))&"</td>" & vbCrLf
      Response.Write "<td title="&rs("MesName")&" nowrap>"&StrLeft(rs("MesName"),39)&"</td>" & vbCrLf
	  if Instr(session("AdminPurview"),"|118,")<>0 then 
		  Response.Write("<td>"&Rs("mobile")&"</td>")
	  else
	    if Instr(session("AdminPurview"),"|120,")<>0 then 
		Response.Write("<td>"&Rs("mobile")&"</td>")
		else
		Response.write("<td>隐藏</td>")
		end if
	  end if
	  if rs("SecretFlag") then
	  	 Response.Write "<td nowrap align='center'><a href='TuiJianMsg.asp?ID="&rs("ID")&"&Action=false'><font color='red'>取 消</font></a></td>" & vbCrLf
	  else
        Response.Write "<td nowrap align='center'><a href='TuiJianMsg.asp?ID="&rs("ID")&"&Action=true'><font color='blue'>推 荐</font></a></td>" & vbCrLf
	  end if
      Response.Write "<td nowrap>"&rs("AddTime")&"</td>" & vbCrLf
      if rs("ViewFlag") then
        Response.Write "<td nowrap align='center'><a href='shenghe.asp?id="&rs("ID")&"&Action=false'><font color='blue'>取 消</font></a></td>" & vbCrLf
      else
        Response.Write "<td nowrap align='center'><a href='shenghe.asp?id="&rs("ID")&"&Action=true'><font color='red'>通 过</font></a></td>" & vbCrLf
	  end if
      Response.Write "<td nowrap>"&rs("ReplyTime")&"</td>" & vbCrLf
      Response.Write "<td width='48' nowrap><a href='MessageEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""审核回复修改留言信息"")'><font color='#330099'>审.回.改</font></a></td>" & vbCrLf
      Response.Write "<td width='14' nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='8' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
	if Instr(session("AdminPurview"),"|304,")<>0 then 
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='删除所选' onClick='ConfirmDel(""您真的要删除这些留言信息吗？"");'></td>" & vbCrLf
end if
  
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='10' nowrap  bgcolor='#EBF2F9'>暂无留言信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='10' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
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

function Guest(ID,Linkman)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From NwebCn_Members where ID="&ID
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    Guest=Linkman
  else
    Guest="<font color='green'>会员：</font><a href='MemEdit.asp?Result=Modify&ID="&ID&"' onClick='changeAdminFlag(""前台会员资料"")'>"&Linkman&"</a>"
  end if
  rs.close
  set rs=nothing
end function 
 Function LookAdd(Sip)
  Dim Str1,Str2,Str3,Str4
  Dim Num
  Dim Irs,sql
  If IsNumeric(Left(sip,2)) Then
    If Sip="127.0.0.1" Then sip="192.168.0.1"
      Str1=Left(Sip,InStr(Sip,".")-1)
      Sip=Mid(Sip,InStr(Sip,".")+1)
      Str2=Left(Sip,InStr(Sip,".")-1)
      Sip=Mid(Sip,InStr(Sip,".")+1)
      Str3=Left(Sip,InStr(Sip,".")-1)
      Str4=Mid(Sip,InStr(Sip,".")+1)
  If IsNumeric(Str1)=0 Or isNumeric(Str2)=0 Or isNumeric(Str3)=0 Or isNumeric(Str4)=0 Then
   Else
    num=CInt(Str1)*256*256*256+CInt(Str2)*256*256+CInt(Str3)*256+CInt(Str4)-1
    Dim adb,aConnStr,AConn
    adb = "../DataBase/ip.mdb"
    aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(adb)
    Set AConn = Server.CreateObject("ADODB.Connection")
    aConn.Open aConnStr
    sql="select country from IPTABLE where StartIPnum <="&num&" and EndIPnum >="&num
    Set irs=AConn.Execute(sql)
    If irs.eof And irs.bof Then 
     LookAdd="中国"
    Else
     Do While Not irs.eof
      LookAdd=LookAdd & Irs(0) 
     Irs.MoveNext
     Loop
    End If
    Irs.Close
    Set Irs=nothing
    Set AConn=Nothing
   End If
  End If
 End Function 
%>