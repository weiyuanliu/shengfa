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
<TITLE>友情连接列表</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script type="text/javascript" src="datepicker/js/jquery.js"></script>
<script src="http://new.cnzz.com/v1/js/datepicker.js" language="JavaScript"></script>
<link href="Images/datepicker.css" rel="stylesheet" type="text/css" />
<style>
 img{border:none;}
</style>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|119,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
On Error Resume Next
dim Result,StartDate,EndDate,Keyword,inputDate
Result=request.QueryString("Result")
StartDate=request.QueryString("st")
EndDate=request.QueryString("et")
Keyword=request.QueryString("Keyword")
'========判断是否具有管理权限
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>广告：添加，修改广告相关的内容</strong></font></td>
  </tr>
  <tr>
    <form name="formSearch" method="post" action="Search.asp?Result=ADV">
          <td bgcolor="#FFFFFF" class="date" >
          
		<script src="http://new.cnzz.com/v1/js/cnzzDatePlugin/WdatePicker.js"></script>
<script src="http://new.cnzz.com/v1/js/cnzzDatePlugin/inputDate.js" language="JavaScript"></script>
<script language="javascript">
function pickerok(){
        $().pickerok();
}

function pickercancel(){
        $().pickercancel();     
}
</script>
<div class="dateinput">
<%
if StartDate="" then
 StartDate = FormatDate(now,11)
end if
if EndDate="" then
 EndDate = FormatDate(now,11)
end if
inputDate = request.QueryString("inputDate")
inputDate = split(inputDate,"至")
if Ubound(inputDate) = 1 then
	StartDate=inputDate(0)
	EndDate=inputDate(1)
end if
%>
<a href="advlist.asp?st=<%=FormatDate(now,11)%>&et=<%=FormatDate(now,11)%>" <%if StartDate = FormatDate(now,11) then%> id="look" <%end if%>>今日</a>
<a href="advlist.asp?st=<%=FormatDate(now-1,11)%>&et=<%=FormatDate(now-1,11)%>" <%if StartDate = FormatDate(now-1,11) then%> id="look" <%end if%>>昨日</a>
<a href="advlist.asp?st=<%=FormatDate(now-7,11)%>&et=<%=FormatDate(now,11)%>" <%if StartDate = FormatDate(now-7,11) then%> id="look" <%end if%>>最近7天</a>
<a href="advlist.asp?st=<%=FormatDate(now-30,11)%>&et=<%=FormatDate(now,11)%>" <%if StartDate = FormatDate(now-30,11) then%> id="look" <%end if%>>最近30天</a>
<a href="advlist.asp?st=<%=year(date)&"-"&month(date)&"-1"%>&et=<%=dateadd("d",-1,dateadd("m",1,year(date)&"-"&month(date)&"-1"))%>" <%if StartDate = year(date)&"-"&month(date)&"-1" and EndDate = dateadd("d",-1,dateadd("m",1,year(date)&"-"&month(date)&"-1")) then%> id="look" <%end if%>>本月</a>
<input style="cursor:pointer;" onClick="window.location.href='advlist.asp?st=<%=CDate(StartDate)-1%>&et=<%=CDate(StartDate)-1%>'" value="前一天" type="button">
<input style="cursor:pointer;" onClick="window.location.href='advlist.asp?st=<%=CDate(StartDate)+1%>&et=<%=CDate(StartDate)+1%>'" value="后一天" <%if CDate(EndDate)+1 > CDate(FormatDate(now,11)) then%>disabled="true"<%end if%> type="button">
&nbsp;&nbsp;&nbsp;&nbsp;选择日期:<span><input id="inputDate" name="inputDate" class="input-one" value="<%=StartDate%>至<%=EndDate%>" size="22" type="text"></span>
<input id="date_search" value="查 询" type="submit">&nbsp;
<input value="2013-01-21" id="headaddstattime" type="hidden">
<input value="<%=FormatDate(now,11)%>" id="headtoday" type="hidden">
</div>		

           
          </td>
        </form>
  </tr>
  <tr>
        <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="advset.asp?Result=Add" onClick='changeAdminFlag("添加广告")'>添加广告</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="advlist.asp" onClick='changeAdminFlag("广告列表")'>查看广告</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form action="DelContent.asp?Result=WAIBU_ADV" method="post" name="formDel" >
    <tr>
      <td width="18" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
      <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>广告名称</strong></font></td>
      <td bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>链接地址</strong></font></td>
      <td width="30" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>IP数</strong></font></td>
      <td width="50" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单数</strong></font></td>
      <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>添加时间</strong></font></td>
      <td colspan="2" width="76" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">      </td>
    </tr>
	<% FriendSiteList() %>
  </form>
</table>
<% if request.QueryString("Result")="ModifySequence" then call ModifySequence() %>
<% if request.QueryString("Result")="SaveSequence" then call SaveSequence() %>
</body>
</html>
<%
'-----------------------------------------------------------
function FriendSiteList()
  dim idCount'记录总数
  dim pages'每页条数
      pages=100
  dim pagec'总页数
  dim page'页码
      page=clng(request("Page"))
  dim pagenc '每页显示的分页页码数量=pagenc*2+1
      pagenc=2
  dim pagenmax '每页显示的分页的最大页码
  dim pagenmin '每页显示的分页的最小页码
  dim datafrom'数据表名
      datafrom="NwebCn_Ads_effect"
  dim datawhere'数据条件
       if Result="Search" then
	  
	     datawhere="where ( ae.ADS_Name like '%" & Keyword &_
		           "%' or ae.ADS_Link like '%" & Keyword &_
		           "%') "
	  else
        datawhere=" "
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
  dim taxis'排序的语句
      taxis="order by ae.id desc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(ID) as idCount from ["& datafrom &"] as ae " & datawhere
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
	if StartDate = EndDate and StartDate<>FormatDate(now,11) then
		'StartDate = Cdate(StartDate)-1
	end if
	if StartDate = EndDate and StartDate=FormatDate(now,11) then
		'EndDate = Cdate(EndDate)+1
	end if
	EndDate = Cdate(EndDate)+1
    sql="select ae.*,(select count(o.id) from NwebCn_order as o where o.ADS_Link = ae.id and (o.addtime between '" & StartDate & "' and '" & EndDate & "') and o.fax=0 ) as ocount,(select count(i.id) from NwebCn_Ip as i where i.adv_id = ae.id and (i.addtime between '" & StartDate & "' and '" & EndDate & "') ) as ip_count from ["& datafrom &"] as ae where ae.id in("& sqlid &") "&taxis
	'response.Write(sql)
    set rs=server.createobject("adodb.recordset")
    rs.open sql,conn,0,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ADS_Name")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ADS_Link")&"</td>" & vbCrLf
'	  if StrLen(rs("ADS_Link"))>53 then
'        Response.Write "<td title="&rs("SiteUrl")&" nowrap>"&StrLeft(rs("ADS_Link"),50)&"</td>" & vbCrLf
'      else
'        Response.Write "<td title="&rs("ADS_Link")&" nowrap>"&rs("ADS_Link")&"</td>" & vbCrLf
'      end if
      Response.Write "<td nowrap>"&rs("ip_count")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ocount")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("AddTime")&"</td>" & vbCrLf
      Response.Write "<td width='48' nowrap><a href='advset.asp?Result=Modify&ID="&rs("ID")&"' onClick='changeAdminFlag(""修改广告"")'><font color='#330099'>修改</font></a></td>" & vbCrLf
 	  Response.Write "<td width='14' nowrap><input name='selectID' type='checkbox' value='"&rs("ID")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='6' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'><input name='submitDelSelect' type='button' class='button'  id='submitDelSelect' value='删除所选' onClick='ConfirmDel(""您真的要删除这些广告吗？"");'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='8' nowrap  bgcolor='#EBF2F9'>暂无广告</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='8' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
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