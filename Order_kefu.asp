<%
if Request.QueryString("l")<>"3DFEED3B7B2697C18AFD1F6625334741" then
%>

<%
else
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="Include/Const.asp" -->
<!--#include file="Include/Conn2.asp" -->
<!--#include file="Include/NoSqlHack.asp" -->
<!--#include file="Include/page.asp" -->
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

<META NAME="Keywords" CONTENT="<% =Keywords %>" />
<META NAME="Description" CONTENT="<% =Descriptions %>" />
<title><%= title & SiteTitle%></title>
<link href="css.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript" src="Script/Html.js"></script>
<script language="javascript" src="Script/flash.js"></SCRIPT>
<script language="javascript" src="Script/jquery-1.8.2.js"></SCRIPT>
<script language="javascript">
	var isMobile=/^1\d{10}$/;   
	$(document).ready(function(){
		$("#cod").click(function(){
		  if($("#Sh_Tel").val().substring(0,1) == 1 || $("#Sh_Tel").val() == '')
			if(!isMobile.test($("#Sh_Tel").val())){
				alert("请检查输入的手机号码是否是11位数!");
				return false;
			}	
		});
		$("#bank").click(function(){
		if($("#On_ShTel").val().substring(0,1) == 1 || $("#On_ShTel").val() == '')
			if(!isMobile.test($("#On_ShTel").val())){
				alert("请检查输入的手机号码是否是11位数!");
				return false;
			}	
		});
		$("#alipay").click(function(){
		if($("#On_ShTel").val().substring(0,1) == 1 || $("#On_ShTel").val() == '')
			if(!isMobile.test($("#On_ShTel").val())){
				alert("请检查输入的手机号码是否是11位数!");
				return false;
			}	
		});
	});
</script>
</head>

<body>
<div style="width:972px; height:50px; margin-left:auto; margin-right:auto;"><img src="images/main.jpg" /></div>
<table width="972" border="0" cellpadding="0" cellspacing="0" class="mag">
  <tr>
    <td height="63" bgcolor="#013976"><table width="972" border="0" align="center" cellpadding="0" cellspacing="0" class="mag">
      <tr>
        <td width="15" height="63"></td>
        <td width="240"><a href="http://www.beloj.com"><img src="images/logo.gif" /></a></td>
        <td width="440" align="center"><img src="images/pic_1.gif" /></td>
        <td width="268"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td align="right" class="cl_fff"><a class="a1" href="addFavorite://behavior" onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('<%=SiteUrl%>');return false;">设为首页</a> | <a class="a1" href="addFavorite://bookmarkit" onClick="return bookmarkit()">加入收藏</a>                                     
            <script language="javascript">
				function bookmarkit(){window.external.addFavorite('<%=SiteUrl%>','<%=SiteTitle%>');return false;}
			</script> </td>
          </tr>
          <tr>
            <td height="36" align="right" valign="bottom"><a href="http://www.beloj.com"><img src="images/pic_2.gif" /></a></td>
          </tr>
        </table></td>
        <td width="9"></td>
      </tr>
    </table></td>
  </tr>
</table>

<div class="bbdiv">
  <div class="drtj" id="dr_cate">
    <table width="972" border="0" align="center" cellpadding="0" cellspacing="0" class="mag">
      <tr>
        <td height="38" bgcolor="#013976" class="cl_fff"><table border="0" cellpadding="0" cellspacing="0" class="mag">
          <tr>
            <td><a href="Index.asp" class="a2">网站首页</a></td>
            <td width="18"></td>
            <td><a href="News.asp" class="a2" target="_blank">新闻中心</a></td>
            <td width="18"></td>
            <td><a href="Products.asp" class="a2"  target="_blank">产品说明</a></td>
            <td width="18"></td>
            <td><a href="Gbook.asp" class="a2"  target="_blank">客户留言</a></td>
            <td width="18"></td>
            <td><a href="FAQ.asp" class="a2"  target="_blank">问题解答</a></td>
            <td width="18"></td>
            <td><a href="Order.asp" class="a2"  target="_blank">在线订购</a></td>
            <td width="18"></td>
            <td><a href="Alipay.asp" class="a2"  target="_blank">支付宝购买</a></td>
            <td width="18"></td>
            <td><a href="Delivery.asp" class="a2"  target="_blank">配送方式</a></td>
            <td width="18"></td>
            <td><a href="Query.asp" class="a2"  target="_blank">发货查询</a></td>
            <td width="18"></td>
            <td><a href="Contact.asp" class="a2"  target="_blank">联系我们</a></td>
          </tr>
        </table></td>
      </tr>
    </table>
  </div>
</div>


<script type="text/javascript"> 
$.fn.smartFloat = function() {
var position = function(element) {
var top = element.position().top, pos = element.css("position");
$(window).scroll(function() {
 
var scrolls = $(this).scrollTop();
if (scrolls > top) {
if (window.XMLHttpRequest) {
element.css({position: "fixed",top: 0});	
} else {
element.css({top:0});	
}
}else {
element.css({
position: pos,
top: $(".bbdiv").offset().top
});	
}
});
};
return $(this).each(function() {
position($(this));						 
});
};
//绑定
$("#dr_cate").smartFloat();
</script>

<%
Dim Action:Action=request("Action")
%>
<table width="972" border="0" cellspacing="0" cellpadding="0" class="mag mg_t10">
  <tr>
    <td width="100%" valign="top">
    
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="100%" height="50" valign="top" background="images/972bg.jpg"><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="10" height="10"></td>
              <td width="141" ></td>
              <td rowspan="2" align="right" class="cl_fff" style="padding-right:15px;"></td>
            </tr>
            <tr>
              <td height="38"></td>
              <td align="center" class="fz_24 fw_bd cl_013974">在线订购</td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td valign="top" class="bk_xb1 bk_zb bk_yb">
          <table width="95%" border="0" cellpadding="0" cellspacing="0" class="mag">
            <tr>
              <td height="15">&nbsp;</td>
            </tr>
            <tr>
              <td class="lh_22">
               <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin:auto;">
                 <%if Action="" then%>
                  <tr>
                    <td width="50%" valign="top" align="center">
                    <table width="450" border="0" cellspacing="0" cellpadding="0" align="center" style="margin-left:auto; margin-right:auto;">
                      <tr>
                        <td width="10"><img src="Images/a3.jpg" width="10" height="43" /></td>
                        <td align="left" class="text4" style="background:url(Images/a4.jpg)">货到付款订购（能够货到付款的朋友请选择此订购方式）</td>
                        <td width="10"><img src="Images/a5.jpg" width="10" height="43" /></td>
                      </tr>
                    </table></td>
                    <td valign="top" align="center">
                    <table width="450" border="0" cellspacing="0" cellpadding="0" align="center" style="margin-left:auto; margin-right:auto;">
                      <tr>
                        <td width="10"><img src="Images/a3.jpg" width="10" height="43" /></td>
                        <td align="left" class="text4" style="background:url(Images/a4.jpg)">银行汇款以及支付宝付款订购（先付款有优惠）</td>
                        <td width="10"><img src="Images/a5.jpg" width="10" height="43" /></td>
                      </tr>
                    </table></td>
                  </tr>
                  <% end if%>
                  <tr>
                    <td colspan="2" align="center">
                      <%if Action="Left" then%>
                        <!--#include file="info/RequestLeft.asp"-->
                      <%elseif Action="Right" then%>
                        <!--#include file="info/RequestRight.asp"-->
                      <%else%>
                          <table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr>
                          <td width="50%">
                          <!--#include file="info/zxdg_left.asp"-->
                          </td>
                          <td>
                          <!--#include file="info/zxdg_right.asp"-->
                          </td>
                          </tr>
                          </table>
                      <%End if%>
                    </td>
                  </tr>
            </table>

           	  </td>
          </tr>
            <tr>
              <td height="15">&nbsp;</td>
            </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>



<table width="973" border="0" cellspacing="0" cellpadding="0" class="mag mg_t20">
  <tr>
    <td height="76" bgcolor="#013976"><table width="960" border="0" cellpadding="0" cellspacing="0" class="mag">
      <tr>
        <td width="248" height="57"><a href="#"><img src="images/logo.gif" /></a></td>
        <td width="655" class="lh_22 cl_fff">Copyright &copy; <%=ComName%> <br /> 地址：<%=Address%><br />
          服务热线：<%=Telephone%>&nbsp;<%=IcpNumber%>&nbsp;<script src="http://s20.cnzz.com/stat.php?id=3312377&web_id=3312377&show=pic" language="JavaScript"></script>

</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

<%
Dim advlink,userip,advlinks

userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If userip = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

lailu = Request.ServerVariables("HTTP_REFERER")

session("advlink")=lailu

advlink = session("advlink")

if advlink = "" then
	advlink = request.ServerVariables("HTTP_HOST")
end if

dim strs:strs=split(advlink,"/")(2)
if request.Cookies("advlink") = 0 then
 dim asql,ars
 set ars=server.CreateObject("adodb.recordset")
 asql="select * from NwebCn_Ads_effect where ADS_Link = '"&advlink&"'"
 ars.open asql,conn,1,3
 if not ars.eof then
     ars("ipcount") = ars("ipcount") + 1
	 ars.update
	 Response.Cookies("advlink") = ars("Id")
	 conn.execute("insert into NwebCn_Ip (adv_id,ip,addtime) Values("&ars("Id")&",'"&userip&"','"&now&"')")
	 else
	  ars.close
	  asql="select * from NwebCn_Ads_effect where ADS_Link = '"&strs&"'"
	  ars.open asql,conn,1,3
	  if not ars.eof then
	    ars("ipcount") = ars("ipcount") + 1
	    ars.update
	    Response.Cookies("advlink") = ars("Id")
		conn.execute("insert into NwebCn_Ip (adv_id,ip,addtime) Values("&ars("Id")&",'"&userip&"','"&now&"')")
	  else
	    Response.Cookies("advlink") = 0
	  end if
 end if
 ars.close
 set rs=nothing
end if
%>
<%
end if
%>