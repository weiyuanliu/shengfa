<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../Include/Const.asp" -->
<!--#include file="../Include/Conn2.asp" -->
<!--#include file="../Include/page.asp" -->
<%
dim rs,sql,SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,IcpNumber,MesViewFlag,syimg,gonggao,ybpz,qq,syjs,otherscount,taobaoid,jobcount,message_note
set rs = server.createobject("adodb.recordset")
sql="select top 1 * from NwebCn_Site"
rs.open sql,conn,1,1
SiteTitle=rs("SiteTitle")
SiteUrl=rs("SiteUrl")
ComName=rs("ComName")
Address=rs("Address")
ZipCode=rs("ZipCode")
Telephone=rs("Telephone")
Fax=rs("Fax")
Email=rs("Email")
Keywords=rs("Keywords")
Descriptions=rs("Descriptions")
IcpNumber=rs("IcpNumber")
MesViewFlag=rs("MesViewFlag")
syimg=rs("syimg")
gonggao=rs("Gonggao")
ybpz=rs("ybpz")
taobaoid=Rs("taobaoid")
otherscount=Rs("otherscount")
QQ=RS("QQ")
jobcount=Rs("jobcount")
syjs=rs("syjs")
message_note=rs("message_note")
rs.close
set rs=nothing '


Function Echo(Str)
 response.Write(Str)&vbcrlf
End Function

Function Or2(Str)
 if len(Str)>0 then
  Or2=Replace(Str,"../","")
  else
  Or2=""
 end if
End Function
Function AboutView(Id)
 Dim rs,sql
 set rs=server.CreateObject("Adodb.recordset")
 sql="Select * from NwebCn_About where ViewFlag=1 and Id = "&Id&""
 rs.open sql,conn,1,1
 if not rs.eof then
   Echo Or2(rs("Content"))
 end if
 rs.close
 set rs=nothing
End Function


Function Guanggao(Id,w,h) '
 Dim rs,sql,Link
 set rs=server.CreateObject("Adodb.recordset")
 sql="Select * from guanggao where viewFlag=1 and Id = "&Id
 rs.open sql,conn,1,1
 if not rs.eof then
  if lcase(right(rs("picture"),3))="swf" then
   echo "<script language=""javascript"" type=""text/javascript"">writeflashhtml(""_swf="& Or2(rs("Picture")) &""", ""_width="& w &""", ""_height="& h &""" ,""_wmode=transparent"");</script>"
  else
	  Link=rs("Link")
	  if Link<>"" then
	   echo "<a href='"&Link&"' target='"&rs("target")&"'>"
	  end if
	   echo "<img src='"&Or2(rs("Picture"))&"' width='"& w &"' height='"& h &"' />"
	  if Link<>"" then
	   echo "</a>"
	  end if
  end if
 end if
 rs.close
 set rs=nothing
End Function 
			Function XXL(X)
				  for i = 1 to X
				  Randomize
				  pass=""
				  Do While Len(pass)<X '随机位数 
				  num1=CStr(Chr((57-48)*rnd+48)) '0~9 
				  pass=pass&num1
				  loop 
				  next
				  XXL=pass
			  End Function
			  
			  Function HaveOrderId(str,X)
			   Dim rs,sql
			   set rs=server.CreateObject("Adodb.recordset")
			   sql="Select * from NwebCn_Order where ProductNo = '"&X&"'"
			   rs.open sql,conn,1,1
			   if rs.eof then
				HaveOrderId=X
				else
				HaveOrderId=HaveOrderId(str,str&right(year(now),1)&month(now)&day(now)&(now)&XXL(5))
			   end if
			   rs.close
			   set rs=nothing
			  End Function
			  
Dim Id,SortId,SortPath,KeyWord
Id=request("Id")
If Id="" or not isnumeric(Id) then Id=0 end if
SortId=request("SortId")
If SortId="" or not isnumeric(SortId) then SortId=0 end if
SortPath=request("SortPath")
KeyWord=request("KeyWord")

	Dim url,fname,F,nm,title
	url=Request.ServerVariables("path_info")   
    fname=mid(url,instrRev(url,"/")+1)   
    F=split(fname,".")
	if fname="" then fname="index.asp" end if
    nm=LCase(F(0))

	select case nm
	  case "index"
		  title=""
	  case "about"
		Dim AboutName,AboutContent
		if id=0 then
		call AboutShow(1)
	  	else
		call AboutShow(Id)
		end if
	  	title=AboutName&" - "
	  case "products"
		   title="产品说明 - "
	  case "productview"
	      title=ProductViewTitle(Id)
	  case "news"
		  title=title & ProductListTitle(SortId,"NwebCn_NewsSort") & "新闻中心 - "
	  case "newsview"
	      title=NewsViewTitle(Id)
	  case "gbook"
		if request.querystring("page") <> 0 then
		  title="客户留言_第"&request.querystring("page")&"页 - "
		else
		  title="客户留言 - "
		end if
	  case "faq"
		  title="问题解答 - "
	  case "order"
		  title="在线订购 - "
	  case "alipay"
		  title="支付宝购买 - "
	  case "delivery"
		  title="配送方式 - "
	  case "query"
		  title="发货查询 - "
	  case "contact"
		  title="联系我们 - "
	end select
	
	Function ProductListTitle(SortId,Table)
	 Dim rs,sql
	 set rs=server.CreateObject("adodb.recordset")
	 sql="select * from "&Table&" where ViewFlag=1 and Id = "&SortId
	 rs.open sql,conn,1,1
	 if not rs.eof then
	   ProductListTitle=rs("SortName")&" - "
	 end if
	 rs.close
	 set rs=nothing
	ENd Function			 
	Function ProductViewTitle(Id)
	 Dim rs,sql
	 set rs=server.CreateObject("adodb.recordset")
	 sql="select * from NwebCn_Products where ViewFlag=1 and Id = "&Id
	 rs.open sql,conn,1,1
	 if not rs.eof then
	   ProductViewTitle=rs("ProductName")&" - "
	 end if
	 rs.close
	 set rs=nothing
	ENd Function	
	Function YsViewTitle(Id)
	 Dim rs,sql
	 set rs=server.CreateObject("adodb.recordset")
	 sql="select * from NwebCn_Others where ViewFlag=1 and Id = "&Id
	 rs.open sql,conn,1,1
	 if not rs.eof then
	   YsViewTitle=rs("OthersName")&" - "
	 end if
	 rs.close
	 set rs=nothing
	ENd Function	
	
	Function NewsViewTitle(Id)
	 Dim rs,sql
	 set rs=server.CreateObject("Adodb.recordset")
	 sql="select * from NwebCn_News where Id="&Id
	 rs.open sql,conn,1,1
	 if not rs.eof then
			NewsViewTitle=rs("NewsName") &" - "
	 end if
	 rs.close
	 set rs=nothing
	ENd Function

%>

<meta http-equiv="Cache-Control" content="no-transform">
<meta name="viewport" content="width=device-width, initial-scale=1.0, minimum-scale=1.0, maximum-scale=1.0, user-scalable=0" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<meta name="apple-mobile-web-app-status-bar-style" content="black" />
<meta name="format-detection" content="telephone=no">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
<META NAME="Keywords" CONTENT="<% =Keywords %>" />
<META NAME="Description" CONTENT="<% =Descriptions %>" />
<title><%= title & SiteTitle%></title>
<link href="style/css/index.css" rel="stylesheet" type="text/css" />
<SCRIPT src="style/js/jQuery132.js" type=text/javascript></SCRIPT>
<SCRIPT src="style/js/lazyload.js" type=text/javascript></SCRIPT>
<link href="css/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div class="body">
  