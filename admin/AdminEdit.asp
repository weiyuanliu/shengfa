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
<TITLE>编辑管理员</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../Include/Md5.asp"-->
<!--#include file="CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|101,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result,Action
Result=request.QueryString("Result")
dim ID,AdminName,Working,Password,vPassword,UserName,Purview,Explain,AddTime,GroupID,GroupName,GroupIdName,groupid1
ID=request.QueryString("ID")
groupid1=request.QueryString("GroupID")
if ID="" then ID=0
call AdminEdit() 
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>网站管理员：添加，修改管理员信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AdminEdit.asp?Result=Add" onClick='changeAdminFlag("添加管理员")'>添加管理员</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="AdminList.asp" onClick='changeAdminFlag("网站管理员")'>查看所有管理员</a></td>    
  </tr>
</table>
<br>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <form name="editForm" method="post" action="AdminEdit.asp?Action=SaveEdit&Result=<%=Result%>&ID=<%=ID%>" onSubmit="return CheckAdminEdit()">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	<table width="100%" border="0" cellpadding="0" cellspacing="0" id=editProduct idth="100%">
	  <tr>
        <td width="120" height="20" align="right">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  
	  <tr <% if ID=1 then Response.Write("style='display:none'")%> >
	    <td height="20" align="right">管理员组：</td>
		<%
		if Result ="Add" then
		%>
		<td><select name="GroupID" class="textfield"><% call SelectGroup() %></select></td>
		<%
		end if
		%>
		<%
		if Result ="Modify" then
		%>
		<td><select name="GroupID" class="textfield"><% call SelectGroup1() %></select></td>
        <%
		end if
		%>
	  </tr>
	  <tr>
        <td height="20" align="right">生　　效：</td>
        <td><input name="Working" type="checkbox" value="1" style="HEIGHT: 13px;WIDTH: 13px;" checked></td>
      </tr>
      <tr>
        <td height="20" align="right">登&nbsp;录&nbsp;名：</td>
        <td><input name="AdminName" type="text" class="textfield" id="AdminName" style="WIDTH: 120;" value="<%=AdminName%>" maxlength="16" <%if Result="Modify" then response.write ("readonly")%>>&nbsp;*&nbsp;3-10位字符，不可修改</td>
      </tr>
      <tr>
        <td height="20" align="right">密　　码：</td>
        <td><input name="Password" type="password" class="textfield" id="Password" maxlength="20" style="WIDTH: 120;">&nbsp;*&nbsp;6-16位字符，不填表未修改密码</td>
      </tr>
      <tr>
        <td height="20" align="right">确认密码：</td>
        <td><input name="vPassword" type="password" class="textfield" id="vPassword" maxlength="20" style="WIDTH: 120;">&nbsp;*</td>
      </tr>
      <tr>
        <td height="20" align="right">管理员名：</td>
        <td><input name="UserName" type="text" class="textfield" id="UserName" style="WIDTH: 120;" value="<%=UserName%>"></td>
      </tr>
      <tr <%if ID=1 then response.write ("style=display:none")%>>
        <td height="20" align="right">操作权限：</td>
        <td>
		<table border="0">
		<tr>
		
		<td>
	      <input name="Purview309" type="checkbox" value="|309," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|309,")>0 then response.write ("checked")%>> 网站信息设置管理
		</td>
		<td style="padding-left:10px">
		 <input name="Purview112" type="checkbox" value="|112," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|112,")>0 then response.write ("checked")%>>&nbsp;网站信息设置
		</td>
		<td style="padding-left:10px">		
		  <input name="Purview310" type="checkbox" value="|310," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|310,")>0 then response.write ("checked")%>> 访问日志管理
		</td>
		<td style="padding-left:10px">
		<input name="Purview312" type="checkbox" value="|312," style="HEIGHT: 13px;WIDTH: 13px;"
		<%if Instr(Purview,"|312,")>0 then response.write ("checked")%>>  回收站管理
		</td>
		<td style="padding-left:10px">
		<input name="Purview105" type="checkbox" value="|105," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|105,")>0 then response.write ("checked")%>>&nbsp;管理员组别		  
		</td>
		<td style="padding-left:10px">
		   <input name="Purview102" type="checkbox" value="|102," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|102,")>0 then response.write ("checked")%>>&nbsp;管理员列表
		 
		</td>
		<td style="padding-left:10px">
		  <input name="Purview101" type="checkbox" value="|101," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|101,")>0 then response.write ("checked")%>>&nbsp;编辑管理员		  
		</td>
		<td style="padding-left:10px">
		  <input name="Purview111" type="checkbox" checked value="|111," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|111,")>0 then response.write ("checked")%>>&nbsp;修改密码<font color="red">*</font>
		</td>
		<td style="padding-left:10px">
		   <input name="Purview104" type="checkbox" value="|104," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|104,")>0 then response.write ("checked")%>>&nbsp;会员列表		 
		</td>
		</tr>
		<tr>
		<td>
		  <input name="Purview311" type="checkbox" value="|311," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|311,")>0 then response.write ("checked")%>>  用户管理权限
          
		</td>
		<td style="padding-left:10px">
		   <input name="Purview103" type="checkbox" value="|103," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|103,")>0 then response.write ("checked")%>>&nbsp;编辑会员
			
		</td>
		<td style="padding-left:10px">
		   <input name="Purview307" type="checkbox" value="|307," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|307,")>0 then response.write ("checked")%>>  修改密码权限
			
		</td>
		<td style="padding-left:10px">
		   <input name="Purview304" type="checkbox" value="|304," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|304,")>0 then response.write ("checked")%>> 留言信息删除
			
		</td>
		<td style="padding-left:10px">
		    <input name="Purview306" type="checkbox" value="|306," style="HEIGHT: 13px;WIDTH: 13px;"
		    <%if Instr(Purview,"|306,")>0 then response.write ("checked")%>> 订单信息删除
		</td>
		<td style="padding-left:10px">
		   <input name="Purview119" type="checkbox" value="|119," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|119,")>0 then response.write ("checked")%>>&nbsp;网络推广
		</td>
		<td style="padding-left:10px">
		   <input name="Purview118" type="checkbox" value="|118," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|118,")>0 then response.write ("checked")%>>&nbsp;显示地区
		</td>
		<td style="padding-left:10px">
		   <input name="Purview120" type="checkbox" value="|120," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|120,")>0 then response.write ("checked")%>>&nbsp;显示IP
		</td>
		<td>
		   <input name="Purview120" type="checkbox" value="|121," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|121,")>0 then response.write ("checked")%>>&nbsp;显示电话号码
		</td>
		</tr>
		<tr>
		<td>
		  <input name="Purview316" type="checkbox" value="|316," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|316,")>0 then response.write ("checked")%>> 订单来源	
		</td>
		<td  style="padding-left:10px">
		  <input name="Purview317" type="checkbox" value="|317," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|317,")>0 then response.write ("checked")%>> 垃圾留言处理		
		</td>
		<td  style="padding-left:10px">
		  <input name="Purview300" type="checkbox" value="|300," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|300,")>0 then response.write ("checked")%>> 企业信息管理		
		</td>
		<td style="padding-left:10px">
		  <input name="Purview11" type="checkbox" value="|11," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|11,")>0 then response.write ("checked")%>>&nbsp;编辑企业
		
		</td>
		<td style="padding-left:10px">		    		
		  <input name="Purview12" type="checkbox" value="|12," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|12,")>0 then response.write ("checked")%>>&nbsp;企业列表
		</td>
		</tr>
		<tr>
		<td>
		   <input name="Purview301" type="checkbox" value="|301," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|301,")>0 then response.write ("checked")%>>  新闻中心管理
		</td>
		<td style="padding-left:10px">
		  <input name="Purview23" type="checkbox" value="|23," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|23,")>0 then response.write ("checked")%>>&nbsp;编辑新闻
		</td>		
		<td style="padding-left:10px">
		  <input name="Purview21" type="checkbox" value="|21," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|21,")>0 then response.write ("checked")%>>&nbsp;新闻类别
		</td>
		<td style="padding-left:10px">
		  <input name="Purview22" type="checkbox" value="|22," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|22,")>0 then response.write ("checked")%>>&nbsp;新闻列表
		</td>				
		<td style="padding-left:10px">
		  
		</td>
		<td style="padding-left:10px">
		  
		</td>
		</tr>
		<tr>
		<td>
		  <input name="Purview302" type="checkbox" value="|302," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|302,")>0 then response.write ("checked")%>>  产品展示管理
		<td style="padding-left:10px">
		  <input name="Purview33" type="checkbox" value="|33," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|33,")>0 then response.write ("checked")%>>&nbsp;编辑产品
		</td>
		<td style="padding-left:10px">
		  <input name="Purview31" type="checkbox" value="|31," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|31,")>0 then response.write ("checked")%>>&nbsp;产品类别	
		</td>
		<td style="padding-left:10px">
		  <input name="Purview32" type="checkbox" value="|32," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|32,")>0 then response.write ("checked")%>>&nbsp;产品列表
		</td>		
		</tr>
		<tr>
		<td>
		   <input name="Purview305" type="checkbox" value="|305," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|305,")>0 then response.write ("checked")%>> 订单信息查看             
		</td>
		<td style="padding-left:10px">
		   <input name="Purview313" type="checkbox" value="|313," style="HEIGHT: 13px;WIDTH: 13px;"            
		   <%if Instr(Purview,"|313,")>0 then response.write ("checked")%>>  订单单独删除
		</td>
		<td style="padding-left:10px">		
		  <input name="Purview314" type="checkbox" value="|314," style="HEIGHT: 13px;WIDTH: 13px;"
          <%if Instr(Purview,"|314,")>0 then response.write ("checked")%>> 订单状态管理
		</td>
		<td style="padding-left:10px">
		   <input name="Purview315" type="checkbox" value="|315," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|315,")>0 then response.write ("checked")%>>&nbsp;订单恢复
		</td>
		<td style="padding-left:10px">
		   <input name="Purview99" type="checkbox" value="|99," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|99,")>0 then response.write ("checked")%>>&nbsp;订单留言列表
		</td>
		<td style="padding-left:10px">
		  <input name="Purview308" type="checkbox" value="|308," style="HEIGHT: 13px;WIDTH: 13px;"
		  <%if Instr(Purview,"|308,")>0 then response.write ("checked")%>>  订单留言信息管理
		</td>
		<td style="padding-left:10px">
		   <input name="Purview93" type="checkbox" value="|93," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|93,")>0 then response.write ("checked")%>>&nbsp;订单列表
		</td>
		<td style="padding-left:10px">
		   <input name="Purview94" type="checkbox" value="|94," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|94,")>0 then response.write ("checked")%>>&nbsp;订单回复
		</td>
		<td>
		   <input name="Purview90" type="checkbox" value="|90," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|90,")>0 then response.write ("checked")%>>&nbsp;订单留言回复查看
		</td>
		<td></td>
		<td></td>		
		</tr>
		<tr>
		<td>
		   <input name="Purview303" type="checkbox" value="|303," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|303,")>0 then response.write ("checked")%>>   留言信息查看回复
		</td>
		<td style="padding-left:10px">
		   <input name="Purview91" type="checkbox" value="|91," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|91,")>0 then response.write ("checked")%>>&nbsp;留言列表
		</td>
		<td style="padding-left:10px">
		   <input name="Purview92" type="checkbox" value="|92," style="HEIGHT: 13px;WIDTH: 13px;"
		   <%if Instr(Purview,"|92,")>0 then response.write ("checked")%>>&nbsp;编辑留言
		</td>
		<td></td>
		</tr>
		</table>
		</td>
      <tr <%if ID<>1 then response.write ("style=display:none")%>>
        <td height="20" align="right">操作权限：</td>
        <td nowrap><font color="#FF0000">内置超级管理员帐号，不可修改！</font></td>
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
</BODY>
</HTML>
<%
sub AdminEdit()
  dim Action,rsCheckAdd,rs,sql
   dim rspur,sqlpur,leftpur
  Action=request.QueryString("Action")
  if Action="SaveEdit" then '保存编辑管理员信息
    set rs = server.createobject("adodb.recordset")	
    if Result="Add" then '创建网站管理员
		
      set rsCheckAdd = conn.execute("select AdminName from NwebCn_Admin where AdminName='" & trim(Request.Form("AdminName")) & "'")
      if not (rsCheckAdd.bof and rsCheckAdd.eof) then '判断此管理员名是否存在
        response.write "<script language=javascript> alert('" & trim(Request.Form("AdminName")) & "管理员已经存在，请换一个登录名再试试！');history.back(-1);</script>"
        response.end
      end if  
	  sql="select * from NwebCn_Admin"
      rs.open sql,conn,1,3
      rs.addnew
      if len(trim(Request.Form("AdminName")))<3 or len(trim(Request.Form("Password")))>10  then
        response.write "<script language=javascript> alert('管理员登录名必填，且字符数为3-10位！');history.back(-1);</script>"
        response.end
      end if	  
      if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>16  then
        response.write "<script language=javascript> alert('管理员密码必填，且字符数为6-16位！');history.back(-1);</script>"
        response.end
      end if
	  if Request.Form("Password")<>Request.Form("vPassword") then 
        response.write "<script language=javascript> alert('两次输入的密码不一样！');history.back(-1);</script>"
        response.end
	  end if
      rs("AdminName")=trim(Request.Form("AdminName"))
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
	  GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("GroupName")=GroupIdName(1)
	  rs("Password")=Md5(Request.Form("Password"))
	  rs("UserName")=trim(Request.Form("UserName"))

	  rs("AdminPurview")=Request.Form("Purview11") & Request.Form("Purview12") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") & Request.Form("Purview23") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") &_
	                     Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") &_
	                     Request.Form("Purview71") & Request.Form("Purview72") & Request.Form("Purview73") &_
	                     Request.Form("Purview81") & Request.Form("Purview82") & Request.Form("Purview97") &_
	                     Request.Form("Purview91") & Request.Form("Purview92") & Request.Form("Purview93") &_
	                     Request.Form("Purview94") & Request.Form("Purview95") & Request.Form("Purview96") &_
	                     Request.Form("Purview98") & Request.Form("Purview99") & Request.Form("Purview101") &_
	                     Request.Form("Purview102") & Request.Form("Purview103") & Request.Form("Purview104") &_
	                     Request.Form("Purview105") & Request.Form("Purview111") & Request.Form("Purview112") &_
	                     Request.Form("Purview113") & Request.Form("Purview114") & Request.Form("Purview115") &_
	                     Request.Form("Purview116") & Request.Form("Purview117") & Request.Form("Purview118") &_
	                     Request.Form("Purview119") & Request.Form("Purview120") & Request.Form("Purview300") &_
			     Request.Form("Purview301") & Request.Form("Purview302") & Request.Form("Purview303") &_
	                     Request.Form("Purview304") & Request.Form("Purview305") & Request.Form("Purview90") &_
                             Request.Form("Purview306") & Request.Form("Purview307") & Request.Form("Purview308") &_
	                     Request.Form("Purview309") & Request.Form("Purview310") & Request.Form("Purview311") &_
	                     Request.Form("Purview312") & Request.Form("Purview313")& Request.Form("Purview314")&_
			     Request.Form("Purview315")&Request.Form("Purview316") & Request.Form("Purview317")
	  rs("Explain")=trim(Request.Form("Explain"))
	  rs("AddTime")=now()
	  
	  
	 
  
   
	end if  
	if Result="Modify" then '修改网站管理员
      sql="select * from NwebCn_Admin where ID="&ID
      rs.open sql,conn,1,3
      rs("AdminName")=trim(Request.Form("AdminName"))
	  'rs("GroupID")=trim(Request.QueryString("GroupID"))
	  'rs("GroupName")=trim(Request.Form("GroupName"))
	  GroupIdName=split(Request.Form("GroupID"),"┎╂┚")
	  rs("GroupID")=GroupIdName(0)
	  rs("GroupName")=GroupIdName(1)
	  if Request.Form("Working")=1 then
        rs("Working")=Request.Form("Working")
	  else
        rs("Working")=0
	  end if
      if trim(Request.Form("Password"))<>"" then
	    if len(trim(Request.Form("Password")))<6 or len(trim(Request.Form("Password")))>20  then
          response.write "<script language=javascript> alert('管理员密码必填，且字符数为6-20位！');history.back(-1);</script>"
          response.end
        end if
	    if Request.Form("Password")<>Request.Form("vPassword") then 
          response.write "<script language=javascript> alert('两次输入的密码不一样！');history.back(-1);</script>"
          response.end
	    end if
	    rs("Password")=Md5(Request.Form("Password"))
	  end if
	  rs("UserName")=trim(Request.Form("UserName"))
	  rs("AdminPurview")=Request.Form("Purview11") & Request.Form("Purview12") &_
	                     Request.Form("Purview21") & Request.Form("Purview22") & Request.Form("Purview23") &_
	                     Request.Form("Purview31") & Request.Form("Purview32") & Request.Form("Purview33") &_
	                     Request.Form("Purview41") & Request.Form("Purview42") & Request.Form("Purview43") &_
	                     Request.Form("Purview51") & Request.Form("Purview52") & Request.Form("Purview53") &_
	                     Request.Form("Purview61") & Request.Form("Purview62") & Request.Form("Purview71") &_
	                     Request.Form("Purview72") & Request.Form("Purview73") & Request.Form("Purview81") &_
	                     Request.Form("Purview82") & Request.Form("Purview90") & Request.Form("Purview91") &_
	                     Request.Form("Purview92") & Request.Form("Purview93") & Request.Form("Purview94") &_
	                     Request.Form("Purview95") & Request.Form("Purview96") & Request.Form("Purview97") &_
	                     Request.Form("Purview98") & Request.Form("Purview99") & Request.Form("Purview101") &_
	                     Request.Form("Purview102") & Request.Form("Purview103") & Request.Form("Purview104") &_
	                     Request.Form("Purview105") & Request.Form("Purview111") & Request.Form("Purview112") &_
	                     Request.Form("Purview113") & Request.Form("Purview114") & Request.Form("Purview115") &_
	                     Request.Form("Purview116") & Request.Form("Purview117") & Request.Form("Purview118") &_
	                     Request.Form("Purview119") & Request.Form("Purview120") & Request.Form("Purview300") &_
			     Request.Form("Purview301") & Request.Form("Purview302") & Request.Form("Purview303") &_
	                     Request.Form("Purview304") & Request.Form("Purview305") & Request.Form("Purview306")&_
	                     Request.Form("Purview307") & Request.Form("Purview308") & Request.Form("Purview309") &_
	                     Request.Form("Purview310") & Request.Form("Purview311")& Request.Form("Purview312") &_
			     Request.Form("Purview313") & Request.Form("Purview314")& Request.Form("Purview315") &_
			     Request.Form("Purview316") & Request.Form("Purview317")
	  rs("Explain")=trim(Request.Form("Explain"))
	end if
	rs.update
	rs.close
    set rs=nothing 
	
	 
    response.write "<script language=javascript> alert('成功编辑网站管理员！');changeAdminFlag('网站管理员');location.replace('AdminList.asp');</script>"
  else '提取管理员信息
	if Result="Modify" then
      set rs = server.createobject("adodb.recordset")
      sql="select * from NwebCn_Admin where ID="& ID
      rs.open sql,conn,1,1
	  AdminName=rs("AdminName")
	  Working=rs("Working")
	  UserName=rs("UserName")
	  Purview=rs("AdminPurview")
	  Explain=rs("Explain")
	  'GroupID=rs("GroupID")
	  GroupID=rs("GroupID")
	  GroupName=rs("GroupName")
	  rs.close
      set rs=nothing 
	end if
  end if
end sub
  
%>
<% 
sub SelectGroup()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from NwebCn_MemGroup where GroupID!=1"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
    response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")  
    'response.write("<option value='"&rs("GroupID")&"'")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext	
  wend
  rs.close
  set rs=nothing
end sub
%>


<% 
sub SelectGroup1()
  dim rs,sql
  set rs = server.createobject("adodb.recordset")
  sql="select GroupID,GroupName from NwebCn_MemGroup where GroupID>1"
  rs.open sql,conn,1,1
  if rs.bof and rs.eof then
    response.write("未设组别")
  end if
  while not rs.eof
	response.write("<option value='"&rs("GroupID")&"┎╂┚"&rs("GroupName")&"'")
    if GroupID=rs("GroupID") then response.write ("selected")
    response.write(">"&rs("GroupName")&"</option>")
    rs.movenext
  wend
  rs.close
  set rs=nothing
end sub
%>