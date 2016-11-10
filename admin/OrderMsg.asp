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
<script language="javascript" src="../Script/Admin.js"></script></HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="select_date.asp"-->
<%
if Instr(session("AdminPurview"),"|99,")=0 then 
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
  <form action="DelContent.asp?Result=MsgData" method="post" name="formDel">
  <tr>
    <td width="56" height="27" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>ID</strong></font></td>
    <td width="145" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>留言人</strong></font></td>
    <td width="194" align="center" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订货时间</strong></font></td>
    <td width="153" align="center" bgcolor="#8DB5E9"><strong><font color="#FFFFFF"><strong>联系电话</strong></font></strong></td>
    <td width="114" align="center" bgcolor="#8DB5E9"><strong><font color="#FFFFFF">状态</font></strong></td>
	<td width="124" align="center" bgcolor="#8DB5E9"><strong><font color="#FFFFFF">留言时间</font></strong></td>
    <td width="141" align="center" bgcolor="#8DB5E9"><strong><font color="#FFFFFF">回复时间</font></strong></td>
	<td nowrap bgcolor="#8DB5E9" align='center'>操作员</td>
    <td width="112" colspan="2" align="center" bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
      <input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSelect" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
      <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">	</td>
  </tr>
  <%call MsgList(20)%>
  </form>
</table>
</BODY>
</HTML>
<%
Sub MsgList(Page_Size)
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from MsgData order by id desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='9'>")
				response.Write("对不起，暂无信息！")
			response.Write("</td>")
		response.Write("</tr>")
	else
		rs.pagesize=page_size
		dim sum_page,total,i
		total=rs.recordcount
		sum_page=total \ page_size
		if total mod page_size <>0 then sum_page=sum_page+1
		dim page
		page=trim(request.querystring("page"))
		if page="" or isnull(page) or (not IsNumeric(page)) then
			page=1
		elseif Cint(Page)<=1 then
			page=1
		elseif Cint(page) => sum_page then
			page=sum_page
		else
			page=Cint(page)
		end if
		rs.absolutepage=page
		
		for i=1 to Page_Size
			if not rs.eof then
				response.Write("<tr bgcolor='#EBF2F9'>")
					response.Write("<td align='center'>")
						response.Write(rs("ID"))
					response.Write("</td>")
					
					response.Write("<td align=center>")
						response.Write(RemoveHTML(rs("Msg_Name")))
					response.Write("</td>")
					
					response.Write("<td>")
						response.Write(rs("Msg_Time"))
					response.Write("</td>")
					
					response.Write("<td>")
						response.Write(RemoveHTML(rs("Msg_TelPhone")))
					response.Write("</td>")
					
					response.Write("<td align=center>")
						if rs("Replay")<>"" then
							response.Write("<font color='#6298E1'>已回复</font>")
						else
							response.Write("<font color='#ff0000'>未回复</font>")
						end if
					response.Write("</td>")
					response.Write("<td>")
						response.Write(rs("Addtime")&"")
					response.Write("</td>")
					response.Write("<td>")
						response.Write(rs("ReplayTime"))
					response.Write("</td>")
					response.Write("<td align='center'>")
						response.Write(rs("Ediadmin"))
					response.Write("</td>")
					response.Write("<td align='center'>")
					response.Write("<a href='RepalyMsg.asp?id="&rs("ID")&"'>")
						if rs("Replay")<>"" then
							response.Write("查看")					
						else
							response.Write("回复")
						end if
					response.Write("</a>")
						response.Write("<input type='checkbox' name='SelectID' id='SelectID' value='"&rs("ID")&"' style='margin-left:10px;'>选择")
						
					response.Write("</td>")
				response.Write("</tr>")
				rs.movenext
			end if
		next
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='8'></td>")
			response.Write("<td align='center'><input type='submit' name='DelRecord' id='DelRecord' value='删 除' onclick='return DelRecords();'></td>")
		response.Write("</tr>")
		if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
	end if
	rs.close()
	set rs=Nothing
End Sub
%>
<%
sub Contrl_Page(page,sum_page,total,page_size) 
dim Url,linkfile,pagewhere,UrlValue
Url=request.ServerVariables("URL")
Url=mid(Url,InstrRev(Url,"/")+1)
linkfile=Url
UrlValue=""
pagewhere=UrlValue

	response.Write("<tr>")
		response.Write("<td colspan='7' class='Item_list' style='padding-top:5px; padding-bottom:5px;'>")
			response.Write("[共计："&total&"条] ")
					response.write("[每页："&page_size&"条] ")
					response.write("[页次："&page&"/"&sum_page&"] ")
					if page<=1 then
						response.write("[首页] [上一页] ")
					else 
						response.write("<a href='"&linkfile&"?page=1"&pagewhere&"'>")
						response.write("[首页]")
						response.write("</a> ")
						response.write("<a href='"&linkfile&"?page="&page-1&pagewhere&"'>")
						response.write("[上一页]")
						response.write("</a> ")
					end if
					
					if page < sum_page then
						response.write("<a href='"&linkfile&"?page="&page+1&pagewhere&"'>")
						response.write("[下一页]")
						response.write("</a> ")
					else
						response.write("[下一页] ")
					end if
					
					if sum_page>1 and page < sum_page then
						response.write("<a href='"&linkfile&"?page="&sum_page&pagewhere&"'>")
						response.write("[末页]")
						response.write("</a>")
					else
						response.write("[末页]")
					end if
					dim cc
					response.write(" 转到：")%>
					<select name="page" size="1" onChange="javascript:window.location='<%=linkfile%>?page='+this.options[this.selectedIndex].value+'<%=pagewhere%>';">
						<%for cc=1 to sum_page
							if cc=page then
								response.write("<option value='"&cc&"' selected >"&cc&"页")
							else
								response.write("<option value='"&cc&"'>"&cc&"页")
							end if
						next%>
					</select>
		<%response.Write("</td>")
	response.Write("</tr>")
end sub
%>

<script language="javascript">
	<!--
	function DelRecords()
	{
		var DelId=document.getElementsByTagName("input");

		var flag=false;
		for(var i=0; i < DelId.length;i++)
		{
			if(DelId[i].type=="checkbox")
			{
				if(DelId[i].status)
				{
					flag=true;
				}
			}
		}
		if(!flag)
		{
			//alert("对不起，你还没选择记录！");
			//return false;
		}
		if(confirm("您确认是否删除所选记录？"))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	-->
</script>