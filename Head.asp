<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="Include/Const.asp" -->
<!--#include file="Include/Conn2.asp" -->
<!--#include file="Include/NoSqlHack.asp" -->
<!--#include file="Include/page.asp" -->
<!--#include file="getwap.asp" -->
<%
call Check_Wap()
%>
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
				  Do While Len(pass)<X '���λ�� 
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
		   title="��Ʒ˵�� - "
	  case "productview"
	      title=ProductViewTitle(Id)
	  case "news"
		  title=title & ProductListTitle(SortId,"NwebCn_NewsSort") & "�������� - "
	  case "newsview"
	      title=NewsViewTitle(Id)
	  case "gbook"
		if request.querystring("page") <> 0 then
		  title="�ͻ�����_��"&request.querystring("page")&"ҳ - "
		else
		  title="�ͻ����� - "
		end if
	  case "faq"
		  title="������ - "
	  case "order"
		  title="���߶��� - "
	  case "alipay"
		  title="֧�������� - "
	  case "delivery"
		  title="���ͷ�ʽ - "
	  case "query"
		  title="������ѯ - "
	  case "contact"
		  title="��ϵ���� - "
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
<title><%= title & SiteTitle%></title>
<META NAME="Keywords" CONTENT="<% =Keywords %>" />
<META NAME="Description" CONTENT="<% =Descriptions %>" />
<link href="style/blue/css/index.css" rel="stylesheet" type="text/css" />
<SCRIPT src="style/blue/js/jQuery132.js" type=text/javascript></SCRIPT>
<SCRIPT src="style/blue/js/lazyload.js" type=text/javascript></SCRIPT>
<SCRIPT type=text/javascript>
$(document).ready(function(){
    $(".loads").find("img").lazyload({effect:"fadeIn",placeholder : "style/blue/images/grey.gif"});
  });
</SCRIPT>
</head>

<body>
<div class="body">
  <div class="body1">
    <div id="header">
     
<div class="htopbg">
       <div class="htop1">
         <div class="logo" ><p style="text-align:center;"><img src="style/blue/images/header_01.jpg" width="1420" height="124" /></p></div>
		
       </div>

<link href="style/blue/css/tip.css" rel="stylesheet" type="text/css" />

<div style="clear:both;"></div>
<script type="text/javascript" charset="utf-8">
$("#thp_notf_div").slideDown();
$(".hpn_top_close").click(function(){
	$("#thp_notf_div").slideUp();
});
</script>

       <div class="menubg">
       <div class="menu">
       <div class="menubox">
         <ul>
          <li<%if nm="index" then%> class="meunh"<%end if%>><a href="Index.asp" title="������ҳ">��ҳ<br><span>HOME</span></a></li>
          <li<%if nm="news" then%> class="meunh"<%end if%>><a href="News.asp" title="��������">��������<br><span>NEWS</span></a></li>
          <li<%if nm="products" then%> class="meunh"<%end if%>><a href="Products.asp" title="��Ʒ˵��">��Ʒ˵��<br><span>PRODUCTS</span></a></li>
          <li<%if nm="faq" then%> class="meunh"<%end if%>><a href="FAQ.asp" title="������">������<br><span>QUESTIONS</span></a></li>
          <li<%if nm="gbook" then%> class="meunh"<%end if%>><a href="Gbook.asp" title="�ͻ�����">�ͻ�����<br><span>FEEDBACK</span></a></li>
          <li<%if nm="delivery" then%> class="meunh"<%end if%>><a href="Delivery.asp" title="���ͷ�ʽ">���ͷ�ʽ<br><span>DELIVERY</span></a></li>
          <li<%if nm="order" then%> class="meunh"<%end if%>><a href="Order.asp" title="���߶���">���߶���<br><span>ORDER</span></a></li>
         
          <li<%if nm="query" then%> class="meunh"<%end if%>><a href="Query.asp" title="������ѯ">������ѯ<br><span>QUERY</span></a></li>
          <li<%if nm="contact" then%> class="meunh"<%end if%>><a href="Contact.asp" ��ϵ����>��ϵ����<br><span>CONTACT</span></a></li>
		  
         </ul>
        </div>
       </div>
       </div>
</div>
<script type="text/javascript">
function show()
{
	window.alert("�����������ϵQQ:1010816714(ע�⣡��QQ��ֻ���ܹ�����Ǣ̸�������κ���������ѯ��ѯ��վ��ؿͷ����߷���绰��)");
}
</script>