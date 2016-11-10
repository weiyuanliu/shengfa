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
<TITLE>订单列表</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<link rel="stylesheet" href="Images/FilesCss.css">
<script language="javascript" src="../Script/Admin.js"></script>
<style type="text/css">
<!--
.STYLE3 {color: #FFFFFF; font-weight: bold; }
-->
</style>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="select_date.asp"-->
<%


'========判断是否具有管理权限
%>
<BODY>
<%
On Error Resume Next
dim Result,StartDate,EndDate,Keyword
Result=request.QueryString("Result")
StartDate=request.QueryString("StartDate")
EndDate=request.QueryString("EndDate")
Keyword=request.QueryString("Keyword")
 
function PlaceFlag()
  dim states
  states=Trim(Request("State"))
  if Result="Search" then
    Response.Write "订单：列表&nbsp;->&nbsp;检索&nbsp;->&nbsp;订购时间[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，关键字[<font color='red'>"&Keyword&"</font>]"
  else
    Response.Write "订单：列表&nbsp;->&nbsp全部"
  end if
  response.Write("<select name='State' id='State' size='1' style='margin-left:10px;' onchange='Evern_Change();'>")
  	response.Write("<option value='NULL'>--全部--</option>")
	
  	if states="未付款" then
		response.Write("<option value='未付款' selected>未付款</option>")
	else
		response.Write("<option value='未付款'>未付款</option>")
	end if
	
	if states="货款已付" then
		response.Write("<option value='货款已付' selected>货款已付</option>")
	else
		response.Write("<option value='货款已付'>货款已付</option>")
	end if
	
	if states="钱到已发" then
		response.Write("<option value='钱到已发' selected>钱到已发</option>")
	else
		response.Write("<option value='钱到已发'>钱到已发</option>")
	end if
	
	if states="不能到付" then
		response.Write("<option value='不能到付' selected>不能到付</option>")
	else
		response.Write("<option value='不能到付'>不能到付</option>")
	end if
	
	if states="已经发货" then
		response.Write("<option value='已经发货' selected>已经发货</option>")
	else	
		response.Write("<option value='已经发货'>已经发货</option>")
	end if
	
	if states="刚订未发" then
		 response.Write("<option value='刚订未发' selected>银行付款</option>")
	else	
		response.Write("<option value='刚订未发'>银行付款</option>")
	end if
	if states="未处理" then
		 response.Write("<option value='未处理' selected>未处理</option>")
	else	
		response.Write("<option value='未处理'>未处理</option>")
	end if
	
	'if states="货到后付款" then
		'response.Write("<option value='货到后付款' selected>货到后付款</option>")
	'else
		'response.Write("<option value='货到后付款'>货到后付款</option>")
	'end if
	'if states="不能发货" then
		'response.Write("<option value='不能发货' selected>不能发货</option>")
	'else
		'response.Write("<option value='不能发货'>不能发货</option>")
	'end if
  response.Write("</select>")
end function  
%>
<script language="javascript">
<!--
function Evern_Change()
{
	var state=document.getElementById("State");
	window.location.href="OrderList.asp?State="+state.value;
}
-->
</script>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" ><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>订单黑名单</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" bgcolor="#EBF2F9">
	<table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="Search.asp?Result=Orderh">
          <td > 订单检索：从
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
          &nbsp;&nbsp;关键字：<input name="Keyword" type="text" class="textfield" value="<%=Keyword%>" size="18" />
          <input name="submitSearch" type="submit" class="button" value="检索" />
          </td>
        </form>
        <td align="right" >查看：<a href="OrderList.asp" onClick='changeAdminFlag("订单信息列表")'>所有订单信息</a></td>
      </tr>
	</table>
	</td>    
  </tr>
</table>

<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"><%PlaceFlag()%><span style="margin-left:20px;"><input type="button" name='Excle' id="Excle" value="导出Excle文件" onClick="Create_ExcelFile();"></span>

	<form name="formxian" method="post" action="BlacklistH.asp" style="margin:0px;display:inline;">
	&nbsp;显示&nbsp;<input name="num" type="text" class="textfield" size="6" maxlength="6" onkeyup="value=value.replace(/[^\d]/g,'')" />&nbsp;条
	<input name="submitxian" type="submit" class="button" value="确定" />
	</form>
</td>
  </tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=OrderD" method="post" name="formDel" >
  <tr>
    
    <td width="106" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订购者</strong></font></td>
    <td width="80" align="center" nowrap bgcolor="#8DB5E9"><span class="STYLE3">拉黑原因</span></td>
    <td width="130" align="center" nowrap bgcolor="#8DB5E9"><span class="STYLE3">订购内容</span></td>
    <td width="108" align="center" bgcolor="#8DB5E9"><span class="STYLE3">联系电话</span></td>
   <td width="108" align="center" nowrap bgcolor="#8DB5E9"><span class="STYLE3">支付方式</span></td>
    <td width="114" align="center"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">订购时间</font></strong></td>
    <td width="62"  align="center" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">状 态</font></strong></td>
    <td colspan="2" width="85"  bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSelect" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">	</td>
	<td nowrap bgcolor="#8DB5E9" align='center'>操作员</td>
  </tr>
  <% OrderList() %>
  </form>
</table>
</BODY>
</HTML>
<%
'-----------------------------------------------------------
	
function OrderList()
On Error Resume Next

if request.Form("num") <> "" then
	session("num") = request.Form("num")
end if
dim num
num=session("num")'接收POST条数
  dim pages'每页条数
if num="" then
      pages=100
else
      pages=num
end if

  dim idCount'记录总数

  dim pagec'总页数
  dim page'页码
      page=clng(request("Page"))
  dim pagenc'每页显示的分页页码数量=pagenc*2+1
      pagenc=2
  dim pagenmax'每页显示的分页的最大页码
  dim pagenmin'每页显示的分页的最小页码
  dim datafrom'数据表名
      datafrom="NwebCn_Order"
  dim datawhere'数据条件
  
      if Result="Search" then
	  
	     datawhere="where ( ProductName like '%" & Keyword &_
		           "%' or ProductNo like '%" & Keyword &_
		           "%' or Linkman like '%" & Keyword &_
		           "%' or Company like '%" & Keyword &_
		           "%') "
	  else
        datawhere=" where Fax<>1 and blacklist=1 "
	  end if
	  if Trim(Request("State"))<>"NULL" and Trim(Request("State"))<>"" then 
	  	if Trim(Request("State"))<>"待处理" then	  
			if datawhere="" then
				datawhere="where State='"&Trim(Request("State"))&"'"
			else
				datawhere=datawhere&" and State='"&Trim(Request("State"))&"'"
			end if
		else
			if datawhere="" then
				datawhere=" where HuoDao_FuKuan=1 and (State is Null)"
			else
				datawhere=datawhere&" and HuoDao_FuKuan=1 and (State is Null)"
			end if
		end if
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
	'response.Write(sql)
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,1,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
    
      Response.Write "<td >"&Guest(rs("MemID"),rs("Linkman"))&"</td>" & vbCrLf
      Response.Write "<td >"&Guest(rs("MemID"),rs("NotSendtext"))&"</td>" & vbCrLf
	  
	  if StrLen(rs("Amount"))>50 then
        Response.Write "<td title="&rs("Amount")&" >"&Replace(replace(Print(rs("Amount")),"倍洛加","(B)"),"二代0盒","")&"</td>" & vbCrLf
      else
        Response.Write "<td title="&rs("Amount")&" >"&Replace(Replace(replace(Print(rs("Amount")),"倍洛加","(B)"),"一代0盒、",""),"、二代0盒","")&"</td>" & vbCrLf
      end if
	  
	  'Response.Write("<td>")
	  'response.Write(rs("Tel"))
	  'response.Write("</td>")
'屏蔽电话号码
dim Telh
'if session("AdminId") = 62 or session("AdminId") = 1 then
    if Instr(session("AdminPurview"),"|121,")>0 then 
	Telh = rs("Tel")
else
	Telh = Left(rs("Tel"),0)&"********"&right(rs("Tel"),3)
end if
	  Response.Write "<td >"&Telh&"</td>" & vbCrLf
	  Response.Write("<td align='center'>"&vbcrlf)
	  	Dim ZiFu_FS
		ZiFu_FS=Split(rs("Remark"),"|")
		Response.Write(ZiFu_FS(1))
		Response.Write(ZiFu_FS(2))
		Response.Write("元")
	  Response.Write("</td>")
	  Response.Write "<td  >"&rs("AddTime")&"</td>" & vbCrLf
	  Response.Write("<td align='center' style='color:#ff0000;'>"&vbcrlf)
	  	'//最新修改的代码
	   if Instr(session("AdminPurview"),"|314,")>0 then
		if rs("State")<>"" and rs("State") <>"未处理" then
			if instr(rs("State"),"货款已付") > 0 then
				response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=钱到已发'>钱到已发</a> | ")
				response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=已经发货'>已经发货</a> | ")
				response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=未处理'>返回状态</a>")
			else
				response.Write(Replace(rs("State"),"刚订未发","银行付款"))
				response.Write("| &nbsp; <a href='ChangState.asp?ID="&rs("ID")&"&State=未处理'>返回状态</a>")
			end if
		else
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=未付款'>未付款</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=不能到付'>不能到付</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=货款已付'>货款已付</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=钱到已发'>钱到已发</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=已经发货'>已经发货</a>| ")
			response.Write("<a href='ChangState.asp?ID="&rs("ID")&"&State=刚订未发'>银行付款</a>")
		end if
	   else
		if rs("State")="" then
	    response.Write("未处理")
		else
	    response.Write(rs("State"))
		end if
	   end if
		'//第二次修改后的代码
		'if rs("State")="未付款" then
			'response.Write("<a href='fukuan.asp?id="&rs("id")&"'><font color='#ff0000'>已付款</font></a>")
		'elseif rs("State")="已付款" then
			'response.Write("<a href='fahuo.asp?id="&rs("id")&"'><font color='#ff0000'>发货</font></a>")
		'elseif rs("State")="货已发" then
			'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
		'elseif rs("State")="未发货" then
			'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
		'elseif rs("State")="货未收到" then
			'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
		'else
			'if rs("HuoDao_FuKuan") then
				'if rs("State")="" or isnull(rs("State")) then
					'response.Write("<a href='HuoDaoFk.asp?id="&rs("ID")&"&Action=true'>货到付款</font>")
					'response.Write("<br />")
					'response.Write("<a href='HuoDaoFk.asp?id="&rs("ID")&"&Action=false'>不能发货</a>")
				'else
					'if rs("FuKuan") then
						'response.Write("<a href='fahuo.asp?id="&rs("ID")&"'>发货</font>")
					'else
						'response.Write("不能发货<br>")
						'response.Write("<a href='HuoDaoFk.asp?id="&rs("ID")&"&Action=true'>改为货到付款</font>")
						
					'end if
				'end if
			'end if
		'end if
	  response.Write("</td>"&vbcrlf) 
	  dim del
	  if Instr(session("AdminPurview"),"|314,")=0 then
	   del=""
	   else
	   del="<a href='huifu.asp?id="&Rs("ID")&"&zt=bl'>恢复订单</a>"
	  end if	
	    
      Response.Write "<td colspan=2 align='center'> "&del&" <a href='OrderEdit.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""查看回复订单信息"")'><font color='#330099'>查看</font></a>&nbsp;"
	  response.Write("&nbsp;")
	  dim X
	  if Instr(session("AdminPurview"),"|313,")=0 then
	   X="disabled"
	  else
	   X=""
	  end if 
	  response.Write("<input name='selectID' "&x&" type='checkbox' value='"&rs("ID")&"' onclick='' style='HEIGHT: 13px;WIDTH: 13px;'>")
      response.Write("</td>")
	  response.Write "<td nowrap align='center'>"&rs("deladmin")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='7'   bgcolor='#EBF2F9'></td>" & vbCrLf

'屏蔽删除回收站
Response.Write "<td colspan='4' align='center'  bgcolor='#EBF2F9'>"
if Instr(session("AdminPurview"),"|304,")<>0 then 
 Response.Write "<input "&x&" name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='删除所选' onClick='ConfirmDel(""本次将彻底删除订单？"");'>"
end if
Response.Write "</td>" & vbCrLf

    Response.Write "</tr>" & vbCrLf
   
  else
    response.write "<tr><td height='50' align='center' colspan='10'   bgcolor='#EBF2F9'>暂无订单信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='10'   bgcolor='#D7E4F7'>" & vbCrLf
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
	  response.write ("[<a href="& myself &"Page="& i &"&State="&Trim(Request("State"))&">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write ("&nbsp;<a href='"& myself &"Page="& page+(pagenc*2+1) &"&State="&Trim(Request("State"))&"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>8</font></a>&nbsp;") '如果页码结束值小于总页数则显示(更后)
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"&State="&Trim(Request("State"))&"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>:</font></a>&nbsp;") '如果页码小于总页数则显示(最后页)	
  '设置分页页码结束===============================
  Response.Write "跳到：第&nbsp;<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('只能在跳转目标页框内输入整数！');this.value='"&Page&"';}"" style='HEIGHT: 18px;WIDTH: 40px;'  type='text' class='textfield' value='"&Page&"'>&nbsp;页" & vbCrLf
  Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick=""GoPage2('"&Myself&"','"&Trim(Request.QueryString("State"))&"')"" value='GO'>" & vbCrLf
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
On Error Resume Next
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
function PringText(Remark)
On Error Resume Next
	dim str,str1,i
	str=split(Remark,"|")
	if ubound(str)>0 then
	str1="送货方式："&str(0)
	str1=str1&vbcrlf
	str1=str1&"支付方式："&str(1)
	str1=str1&vbcrlf
	str1=str1&"应付金额："&str(2)
	PringText=str1
	end if
end function

function Print(Amount)
On Error Resume Next
	dim str,i,str1,aa
	str1=""
	aa=replace(Amount,"（","(")
	aa=replace(aa,"）",")")
	str=split(aa,"|")
	'if ubound(str)>0 then
	for i=0 to ubound(str)
		if i>0 then str1=str1&"、"
		if str1="" then
			str1=Mid(str(i),1,instr(str(i),"(")-1)
		else
			str1=str1&Mid(str(i),1,instr(str(i),"(")-1)
		end if
		str1=str1&Mid(str(i),instr(str(i),"(")+1,(instr(str(i),")"))-(instr(str(i),"(")+1))&"盒"
	next
	Print=str1
	'else
	'Print=Amount
	'end if
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
    adb = "../DATAbase/ip.mdb"
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
<script language="javascript">
<!--
function GoPage2(Myself,str2)
{
	window.location.href=Myself+"Page="+document.formDel.SkipPage.value+"&State="+str2;
}
function Create_ExcelFile()
{	
	var iframe=document.createElement("iframe");
	iframe.style.width=0;
	iframe.style.height=0;
	iframe.style.border=0;
	
	var UrlValue=window.location.href;
	var GetValue;
	UrlValue=UrlValue.slice(UrlValue.lastIndexOf("/")+1,UrlValue.length);
	if(UrlValue.indexOf("?")!=-1)
	{
		GetValue=UrlValue.slice(UrlValue.indexOf("?")+1,UrlValue.length)
	} 
	iframe.src="CreateExcle.asp?"+GetValue;
	window.document.body.appendChild(iframe);
	
}

-->

</script>