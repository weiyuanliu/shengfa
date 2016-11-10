<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>古那迪GULEND中国官方网站(中国地区古那迪唯一官方网站）GULEND.COM</title>
<META NAME="Keywords" CONTENT="古那迪GULEND" />
<META NAME="Description" CONTENT="古那迪GULEND中国地区官方网站" />
<link href="style/blue/css/index.css" rel="stylesheet" type="text/css" />
<SCRIPT src="style/blue/js/jQuery132.js" type=text/javascript></SCRIPT>
<SCRIPT src="style/blue/js/lazyload.js" type=text/javascript></SCRIPT>
<SCRIPT type=text/javascript>
$(document).ready(function(){
    $(".loads").find("img").lazyload({effect:"fadeIn",placeholder : "style/blue/images/grey.gif"});
  });
</SCRIPT>
</head>
<!--#include file="Include/NoSqlHack.asp" -->
<!--#include file="Include/Const.asp" -->
<!--#include file="Include/Conn2.asp" -->
<body>
<div class="body">
  <div class="body1">
    <div id="header">
     
<div class="htopbg">
       <div class="htop1">
         <div class="logo"><img src="style/blue/images/header_01.jpg" width="1420" height="131" /></div>
       </div>
       <div class="menubg">
       <div class="menu">
       <div class="menubox">
         <ul>
          <li class="meunh"><a href="Index.asp" title="官网首页">首页<br><span>HOME</span></a></li>
          <li class=""><a href="News.asp" title="新闻中心">新闻中心<br><span>NEWS</span></a></li>
          <li class=""><a href="Products.asp" title="产品说明">产品说明<br><span>PRODUCTS</span></a></li>
          <li class=""><a href="FAQ.asp" title="问题解答">问题解答<br><span>QUESTIONS</span></a></li>
          <li class=""><a href="Gbook.asp" title="客户留言">客户留言<br><span>FEEDBACK</span></a></li>
          <li class=""><a href="Delivery.asp" title="配送方式">配送方式<br><span>DELIVERY</span></a></li>
          <li class=""><a href="Order.asp" title="在线订购">在线订购<br><span>ORDER</span></a></li>
          <li class=""><a href="Alipay.asp" title="支付宝购买">支付宝购买<br><span>ALIPAY</span></a></li>
          <li class=""><a href="Query.asp" title="发货查询">发货查询<br><span>QUERY</span></a></li>
          <li style=" margin-right:0px;"  class=""><a href="Contact.asp" 联系我们>联系我们<br><span>CONTACT</span></a></li>
         </ul>
        </div>
       </div>
       </div>
</div>
     <div style="background:url(style/blue/images/header_03.jpg) center  no-repeat; width:1420px;height:334px;margin:0 auto;"></div>
     <div style="background:url(style/blue/images/header_05.jpg) center  no-repeat; width:1420px;height:111px;margin:0 auto;"></div>
 </div>
<style>
.tijiao{
    cursor:pointer;
    color: rgba(255,255,255,1);
    text-decoration: none;
    background-color: rgba(219,87,5,1);
    font-family: 'Yanone Kaffeesatz';
    font-weight: 700;
    font-size: 14px;
    padding: 4px;
    -webkit-border-radius: 8px;
    -moz-border-radius: 8px;
    border-radius: 8px;
    border: 0;
	width: 80px;
	text-align: center;
	-webkit-transition: all .1s ease;
	-moz-transition: all .1s ease;
	-ms-transition: all .1s ease;
	-o-transition: all .1s ease;
	transition: all .1s ease;
}
</style>
  <div id="main">
    <div class="html">
    <div class="html1">       
<div class="order_ok" >   
<TABLE  cellSpacing="1" cellpadding="1" width="90%"  align=center >
  <tr><td colspan='2' align="center"><img src="style/blue/images/order_ok.jpg" width="552" height="93"></td></tr>
<tr><td colspan="2" class="orderok_t">

<%
dim rs,sql,SiteTitle,SiteUrl,ComName,Address,ZipCode,Telephone,Fax,Email,Keywords,Descriptions,IcpNumber,MesViewFlag,syimg,gonggao,ybpz,qq,syjs
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
QQ=RS("QQ")
syjs=rs("syjs")
rs.close
set rs=nothing 
'
%>

<%
call savedingdan()
%>
<%
function savedingdan()
	Dim ProdName,dgtime,AddTime,Sh_Name,Sh_Mobel,Sh_Tel,Products,AddresStr,fangshi,zifu,ZipCodes,ProductNo,ipadd,Sh_Telto,On_ipto
    ProdName=Trim(Request.form("ProdName"))
    dgtime=Trim(request.Form("dgtime"))
	Sh_Name=Trim(Request.Form("Sh_Name"))
	'Sh_Mobel=Trim(Request.Form("Sh_Mobel"))
	Sh_Tel=Trim(Request.Form("Sh_Tel"))
	Sh_Telto=Trim(Request.Form("Sh_Telto"))
	Products=Trim(Request.Form("Products"))
	AddresStr=Trim(Request.Form("AddresStr"))
	fangshi=Trim(Request.Form("fangshi"))
	zifu=Trim(Request.Form("zifu"))
	ZipCodes=Trim(Request.Form("ZipCode"))
	ProductNo=Trim(Request.Form("ProductNo"))
	AddTime=Trim(Request.Form("dgtime"))
	ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR")
	On_ipto=Trim(Request.Form("On_ipto"))

	dim rs,sql
		set rs=server.CreateObject("adodb.recordset")
	sql="select * from NwebCn_Order where ProductNo='"&ProductNo&"'"

	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('不能重复提交定单！');")
			response.Write("window.history.go(-1);")
		response.Write("</script>")
		response.End()
		exit function
	end if
	
	rs.close()
	Dim Remark,ADS_Link
	ADS_Link = request.Cookies("advlink")

	Remark=fangshi&"|"&zifu&"|"&Trim(Request.form("SumMemony"))

	'查询是否在黑名单内
	dim rsbl,sqlbl,blacklist
	set rsbl=server.CreateObject("adodb.recordset")
	sqlbl="select * from NwebCn_Order where Tel='"&Sh_Tel&"' and Fax='0' and blacklist='1'"
	rsbl.open sqlbl,conn,1,1
	if not rsbl.eof and not rsbl.bof then
	blacklist="2"
	else
	blacklist="0"
	end if
	rsbl.close()

	sql="insert into NwebCn_Order (ProductName,AddTime,Linkman,Address,ZipCode,Telephone,Amount,ProductNo,tel,telto,ipaddress,ipto,Remark,ADS_Link,blacklist) VALUES('"&ProdName&"','"&AddTime&"','"&Sh_Name&"','"&AddresStr&"','"&ZipCodes&"','"&Sh_Mobel&"','"&Products&"','"&ProductNo&"','"&Sh_Tel&"','"&Sh_Telto&"','"&ipadd&"','"&On_ipto&"','"&Remark&"',"&ADS_Link&","&blacklist&")"


	conn.execute(sql)
	if instr(zifu,"支付宝")>0 then
		Dim id,subject,body,order_id,Memony,product_count,yinfei,Key
		id="2088402551356533"
		Key="dinrqtpwtcai6wzv4iy8qby016hb67uo"
		subject=ProdName
		product_count=1
		body=ProdName
		order_id=ProductNo
		'if Sh_Name="唐彬测试" then
		'Memony="0.01"
		'else
		Memony=Trim(Request.form("SumMemony"))
		'end if
		yinfei=0
		call sendSms(5,Sh_Name,Sh_Tel) 
		response.Redirect("new_asp/index.asp?id="&id&"&Key="&Key&"&subject="&subject&"&product_count="&product_count&"&body="&body&"&order_id="&order_id&"&Memony="&Memony&"&yinfei="&yinfei&"&return_url="&SiteUrl&"/OnPlay.asp"&"&seller_email="&Email)	
			
	elseif instr(zifu,"网银")>0 then
		Dim v_mid,keys,v_oid,v_amount
		dim v_rcvname,v_rcvaddr,v_rcvtel,v_rcvpost,v_rcvemail,v_rcvmobile
		dim v_ordername,v_orderaddr,v_ordertel,v_orderpost,v_orderemail,v_ordermobile,remark1,remark2
		
		v_mid=Get_Value("NwebCn_Site","WY_ID")
		keys=Get_Value("NwebCn_Site","WY_Key")
		v_oid=ProductNo
		v_amount=Csng(Trim(Request.form("SumMemony")))
		
		v_rcvname=Sh_Name
		v_rcvaddr=AddresStr
		v_rcvtel=Sh_Tel
		v_rcvpost=ZipCodes
		v_rcvmobile=Sh_Mobel
		call sendSms(3,Sh_Name,Sh_Tel) 
		response.Redirect("wangyun/Send.asp?v_mid="&v_mid&"&key="&keys&"&v_oid="&v_oid&"&v_amount="&v_amount&"&v_rcvname="&v_rcvname&"&v_rcvaddr="&v_rcvaddr&"&v_rcvtel="&v_rcvtel&"&v_rcvpost="&v_rcvpost&"&v_rcvmobile="&v_rcvmobile)	
	else
		'【银行和支付宝】
		call sendSms(3,Sh_Name,Sh_Tel) 
		Call OK()
	end if
end function

function Get_Value(tablename,ziduan)
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select top 1 "&ziduan&" from "&tablename
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		Get_Value=""
	else
		Get_Value=rs(ziduan)
	end if
	rs.close()
	set rs=Nothing
end function

Sub Ok()
'货到付款
response.Write(GetValues("NwebCn_About","Content",56))
End sub

%>
</td></tr>
   <TR>
	<TD colspan="2" align=center  style=" padding-top:10px"><input type="button" name="getbak" class="tijiao" value="返 回" onclick="window.location.href='Index.asp';"></TD>
  </TR>
    <TR>
    	<TD colspan='2' class="ordersm"></TD>
    </TR>
</TABLE>
</div>	

    </div>
   <div>
  </div>
  </div>
 </div><!--#Include file="Order_Foot.asp"-->