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
<TITLE>管理员列表</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/ThreeLd.js"></script>
</HEAD>
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="CheckAdmin.asp"-->
<!--#include file="TreeLDClass.asp"-->
<%
if Instr(session("AdminPurview"),"|82,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
Dim LianDong,Action,ID
Action=Trim(Request.QueryString("Action"))
ID=Trim(Request.QueryString("ID"))
if id="" or isnull(id) or Not(IsNumeric(id)) then
	Call Message("数据出错，请返回！")
	response.End()
end if

if Action="EditRecord" then Call EditRecords()
Set LianDong=New LdClass
LianDong.Set_ID(ID)

'定义全局变量用于保存数据库中的值 
Dim QY_Names,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_XingZhi
Dim QY_Wai,QY_CaoZuo,QY_BeiZu,QY_AddTime,QY_Px,QY_FanWei
Call FuZhi()
LianDong.Set_QY_ShengFen(QY_ShengFen)
LianDong.Set_QY_City(QY_City)
LianDong.Set_QY_Citys(QY_Citys)
%>
<BODY>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#6298E1">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="Images/Explain.gif" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>区域信息添加</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="AddRegional.asp?Result=Add" onClick='changeAdminFlag("添加信息")'>添加信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="Regionallist.asp" onClick='changeAdminFlag("查看信息")'>查看信息</a></td>
  </tr>
  <tr>
    <td height="48" align="center" nowrap  bgcolor="#EBF2F9" style="padding:10px;">
    
    <table width="87%" border="0" cellpadding="4" cellspacing="0">
     <form name="AddRegional" id="AddRegInonal" action="EditRegiona.asp?Action=EditRecord&ID=<%=ID%>" method="post" onSubmit="return Check_AddRegionalValues();">
      <tr>
        <td width="17%" align="right"><strong>区域名称：</strong></td>
        <td width="83%"><label>
          <input name="QY_Names" type="text" id="QY_Names" value="<%=QY_Names%>">　
          *必填
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>省份选择：</strong></td>
        <td><label>
          <select name="QY_ShengFen" id="QY_ShengFen" onChange="ChangEvent('QY_ShengFen','QY_City','QY_Citys','GetLDValue.asp?Action=Two');">
          	<%=LianDong.FirstGread%>
          </select>
        　*必选</label></td>
      </tr>
      <tr>
        <td align="right"><strong>市级选择：</strong></td>
        <td><label>
          <select name="QY_City" id="QY_City" onChange="ChangEvent('QY_City','QY_Citys','Null','GetLDValue.asp?Action=Three');">
          	<%=LianDong.TwoGread%>
          </select>
        　*必选</label></td>
      </tr>
      <tr>
        <td align="right"><strong>区域选择：</strong></td>
        <td><label>
          <select name="QY_Citys" id="QY_Citys">
          	<%=LianDong.ThreeGread%>
          </select>
        　*必选</label></td>
      </tr>
      <tr>
        <td align="right"><strong>网点类型：</strong></td>
        <td><label>
          <input name="QY_Type" type="text" id="QY_Type" value="<%=QY_Type%>">
        　
          *必填
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>网点性质：</strong></td>
        <td><label>
          <input name="QY_XingZhi" type="text" id="QY_XingZhi" value="<%=QY_XingZhi%>">
        　
          *必填
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>服务区域：</strong></td>
        <td>
            <INPUT type="hidden" name="QY_FanWei" value="<%=QY_FanWei%>">
            <IFRAME ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=QY_FanWei&style=s_mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME>       
       </td>
      </tr>
      <tr>
        <td align="right"><strong>服务区域外：</strong></td>
        <td>
        	<INPUT type="hidden" name="QY_Wai" value="<%=QY_Wai%>">
            <IFRAME ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=QY_Wai&style=s_mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME> 
        </td>
      </tr>
      <tr>
        <td align="right"><strong>可操作：</strong></td>
        <td><label>
          <input type="radio" name="QY_CaoZuo" id="QY_CaoZuo" value="１" <%if QY_CaoZuo then response.Write("checked")%>>
          是
         　 
         <input type="radio" name="QY_CaoZuo" id="QY_CaoZuo2" value="０" <%if Not(QY_CaoZuo) then response.Write("checked")%>>
          否
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>备注：</strong></td>
        <td>
        	<INPUT type="hidden" name="QY_BeiZu" value="<%=QY_BeiZu%>">
            <IFRAME ID="eWebEditor1" src="../eWebEditor/ewebeditor.asp?id=QY_BeiZu&style=s_mini" frameborder="0" scrolling="no" width="100%" height="150"></IFRAME> 
        </td>
      </tr>
      <tr>
        <td align="right"><strong>添加时间：</strong></td>
        <td><label>
          <input name="QY_AddTime" type="text" id="QY_AddTime" value="<%=QY_AddTime%>">
        </label></td>
      </tr>
      <tr>
        <td align="right"><strong>排序顺序：</strong></td>
        <td><label>
          <input name="QY_Px" type="text" id="QY_Px" size="10" value="<%=QY_Px%>">
        　请填写数字排序信息，值越大排在越前</label></td>
      </tr>
      <tr>
        <td align="right">&nbsp;</td>
        <td><label>
          <input type="submit" name="tijiao" id="tijiao" value="修 改" style="margin-left:15px; margin-right:10px;">
          <input type="button" name="GetBak" id="GetBak" value="返 回" onClick="window.history.go(-1);">
        </label></td>
      </tr>
      </form>
    </table>    </td>    
  </tr>
</table>
<br>
</body>
</html>
<%
Sub Regionallist(Page_Size)
	Dim Rs,Sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select QY_Names,ID,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_AddTime,QY_Px from Regional order by id asc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='9'>"&vbcrlf)
				response.Write("暂无信息！")
			response.Write("</td>"&vbcrlf)
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
		
		for i=1 to page_size
			if not rs.eof then
				response.Write("<tr bgcolor='#EBF2F9'>")
					response.Write("<td align='center'>")
						response.Write(rs("ID"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(Get_Values("Province","Content",rs("QY_ShengFen")))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(Get_Values("City","Content",rs("QY_City")))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(Get_Values("County","Content",rs("QY_Citys")))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_Names"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_Type"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_Px"))
					response.Write("</td>")
					
					response.Write("<td align='center'>")
						response.Write(rs("QY_AddTime"))
					response.Write("</td>")
				response.Write("</tr>")
				rs.movenext
			end if
		next
		
		response.Write("<tr bgcolor='#EBF2F9'>")
			response.Write("<td colspan='8'></td>")
			response.Write("<td align='center'>")
				response.Write("<input name='DelRecord' type='submit' value='删 除'>")
			response.Write("</td>")
		response.Write("</tr>")
		
		if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
	end if
End sub
%>
<%
sub Contrl_Page(page,sum_page,total,page_size) 
dim Url,linkfile,pagewhere,UrlValue
Url=request.ServerVariables("URL")
Url=mid(Url,InstrRev(Url,"/")+1)
linkfile=Url
UrlValue=""

if UrlValue<>"" then
	pagewhere=UrlValue
end if

	response.Write("<tr bgcolor='#EBF2F9'>")
		response.Write("<td colspan='9' class='Item_list' style='padding-top:5px; padding-bottom:5px; text-align:right;'>")
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

function Get_Values(tablename,Content,ID)
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select "&Content&" from "&tablename&" where id="&ParentID
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		Get_Values=rs(Content)		
	end if
	rs.close()
	set rs=Nothing
end function

Sub EditRecords()
	Dim QY_Names,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_XingZhi
	Dim QY_Wai,QY_CaoZuo,QY_BeiZu,QY_AddTime,QY_Px,QY_FanWei
	
	QY_Names=Trim(Request.Form("QY_Names"))
	QY_ShengFen=Trim(Request.Form("QY_ShengFen"))
	QY_City=Trim(Request.Form("QY_City"))
	QY_Citys=Trim(Request.Form("QY_Citys"))
	QY_Type=Trim(Request.Form("QY_Type"))
	QY_XingZhi=Trim(Request.Form("QY_XingZhi"))
	
	QY_Wai=Trim(Request.Form("QY_Wai"))
	QY_CaoZuo=Trim(Request.Form("QY_CaoZuo"))
	QY_BeiZu=Trim(Request.Form("QY_BeiZu"))
	QY_AddTime=Trim(Request.Form("QY_AddTime"))
	QY_Px=Trim(Request.Form("QY_Px"))
	QY_FanWei=Trim(Request.Form("QY_FanWei"))
	
	if QY_Names="" or isnull(QY_Names) then
		Call Message("请填写名字！")
		response.End()	
	end if
	
	if QY_ShengFen="" or isnull(QY_ShengFen) or QY_ShengFen="Null" or not(IsNumeric(QY_ShengFen)) then
		Call Message("数据不能为空，请返回！")	
		response.End()
	end if
	
	if QY_City="" or isnull(QY_City) or QY_City="Null" or not(IsNumeric(QY_City)) then
		Call Message("数据不能为空，请返回！")	
		response.End()
	end if
	
	if QY_Citys="" or isnull(QY_Citys) or QY_Citys="Null" or not(IsNumeric(QY_Citys)) then
		Call Message("数据不能为空，请返回！")	
		response.End()
	end if
	
	if QY_Type="" or isnull(QY_Type) then
		Call Message("数据不能为空，请返回！")	
		response.End()
	end if
	
	if QY_XingZhi="" or isnull(QY_XingZhi) then
		Call Message("数据不能为空，请返回！")	
		response.End()
	end if
	
	if QY_Px="" or isnull(QY_Px) or Not(IsNumeric(QY_Px)) then
		Call Message("数据出错，请返回！")	
		response.End()
	end if
	
	if QY_FanWei="" or isnull(QY_FanWei) then
		Call Message("数据不能为空，请返回！")	
		response.End()
	end if
	
	Dim Rs,Sql
	Set Rs=Server.CreateObject("adodb.recordset")
	Sql="Select * from Regional where QY_Names='"&QY_Names&"' and QY_ShengFen="&QY_ShengFen&" and QY_City="&QY_City&" and QY_Citys="&QY_Citys&" and id not in("&ID&")"
	Rs.open Sql,conn,1,1
	if Not rs.eof and Not rs.bof then
		rs.close()
		set rs=Nothing
		Call Message("对不起，此记录已经存在，请返回！")
		response.End()
		exit sub
	end if
	rs.close
	Sql="Select top 1 * from Regional Where ID="&ID
	Rs.open sql,conn,1,3
		rs("QY_Names")=QY_Names
		rs("QY_ShengFen")=QY_ShengFen
		rs("QY_City")=QY_City
		rs("QY_Citys")=QY_Citys
		rs("QY_Type")=QY_Type
		rs("QY_XingZhi")=QY_XingZhi
		rs("QY_FanWei")=QY_FanWei
		rs("QY_Wai")=QY_Wai
		if QY_CaoZuo then
			rs("QY_CaoZuo")=true
		else
			rs("QY_CaoZuo")=false
		end if
		rs("QY_BeiZu")=QY_BeiZu
		if QY_AddTime<>"" then
			rs("QY_AddTime")=QY_AddTime
		else
			rs("QY_AddTime")=Now()
		end if	
		rs("QY_Px")=QY_Px
	rs.update()
	rs.close()
	set rs=Nothing
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('记录修改成功！');"&vbcrlf)
		response.Write("window.location.href='Regionallist.asp';")
	response.Write("</script>"&vbcrlf)
End Sub

Sub Message(str)
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('"&str&"');")&vbcrlf
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>"&vbcrlf)
End Sub

Sub FuZhi() '用于读取数据库的某条记录的值，并保存在全局变量中
	Dim Rs,Sql
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	Sql="Select * from Regional where id="&ID
	Rs.Open Sql,conn,1,1
	if rs.eof and rs.bof then
		rs.close()
		set rs=Nothing
		Call Message("对不起，记录未找到！")
		response.End()
		exit sub
	else
		QY_Names=Rs("QY_Names")
		QY_ShengFen=Rs("QY_ShengFen")
		QY_City=Rs("QY_City")
		QY_Citys=rs("QY_Citys")
		QY_Type=rs("QY_Type")
		QY_XingZhi=rs("QY_XingZhi")
		
		QY_Wai=rs("QY_Wai")
		QY_CaoZuo=rs("QY_CaoZuo")
		QY_BeiZu=rs("QY_BeiZu")
		QY_AddTime=rs("QY_AddTime")
		QY_Px=rs("QY_Px")
		QY_FanWei=rs("QY_FanWei")
	end if
	rs.close()
	set rs=Nothing
End Sub
%>

