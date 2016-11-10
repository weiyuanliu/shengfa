<% Option Explicit %>
<html>
<head>
<style type="text/css">
<!--
body {
	background-color: #2A0000;
}
-->
</style>
<% response.charset="gb2312" %>
<!--#include file="Include/NoSqlHack.asp" -->
<!--#include file="Include/Const.asp" -->
<!--#include file="Include/Conn2.asp" -->
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
set rs=nothing '
%>
</head>
<center>
<body>
<%=savedingdan()%>
<%
function savedingdan()
	Dim ProdName,dgtime,AddTime,Sh_Name,Sh_Mobel,Sh_Tel,Products,AddresStr,fangshi,zifu,ZipCodes,ProductNo,ipadd
    ProdName=Trim(Request.form("ProdName"))
    dgtime=Trim(request.Form("dgtime"))
	Sh_Name=Trim(Request.Form("Sh_Name"))
	'Sh_Mobel=Trim(Request.Form("Sh_Mobel"))
	Sh_Tel=Trim(Request.Form("Sh_Tel"))
	Products=Trim(Request.Form("Products"))
	AddresStr=Trim(Request.Form("AddresStr"))
	fangshi=Trim(Request.Form("fangshi"))
	zifu=Trim(Request.Form("zifu"))
	ZipCodes=Trim(Request.Form("ZipCode"))
	ProductNo=Trim(Request.Form("ProductNo"))
	AddTime=Trim(Request.Form("dgtime"))
	ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR") 
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
	sql="select top 1 * from NwebCn_Order"
	rs.open sql,conn,1,3
	rs.addnew()
	rs("ProductName")=ProdName
	rs("AddTime")=AddTime
	rs("Linkman")=Sh_Name
	rs("Address")=AddresStr
	rs("ZipCode")=ZipCodes
	rs("Telephone")=Sh_Mobel
	rs("Amount")=Products
	rs("ProductNo")=ProductNo
	rs("tel")=Sh_Tel
	
	rs("Remark")=fangshi&"|"&zifu&"|"&Trim(Request.form("SumMemony"))
	rs.update()
	rs.close()
	set rs=Nothing
	if instr(zifu,"支付宝")>0 then
		Dim id,subject,body,order_id,Memony,product_count,yinfei,Key
		id="2088402010825864"
		Key="bx9ntzn26lm79u5fem9qfq3flo5cmizf"
		subject=ProdName
		product_count=1
		body=ProdName
		order_id=ProductNo
		
		Memony=Trim(Request.form("SumMemony"))
		yinfei=0
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
		response.Redirect("wangyun/Send.asp?v_mid="&v_mid&"&key="&keys&"&v_oid="&v_oid&"&v_amount="&v_amount&"&v_rcvname="&v_rcvname&"&v_rcvaddr="&v_rcvaddr&"&v_rcvtel="&v_rcvtel&"&v_rcvpost="&v_rcvpost&"&v_rcvmobile="&v_rcvmobile)	
	else
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
	response.Write("<div style='width:400px; height:300px; background:#ffffff; padding:10px; border:#000000 1px solid;'>")
		response.Write("<div style='text-align:center; font-size:18px; color:#ff0000; padding-top:10px; padding-bottom:10px;'><strong>订单已经提交成功！！</strong></div>")
		response.Write("<div style='text-align:left; line-height:25px;'>")
			response.Write(GetValues("NwebCn_About","Content",56))
		response.Write("</div>")
		response.Write("<div style='padding-top:10px;'>")
			response.Write("<input type='button' name='getbak' id='getbak' value='返 回' style='border:#ff0000 1px solid; font-size:14px; padding-top:3px;' onclick=""window.location.href='default.asp';"">")
		response.Write("</div>")
	response.Write("</div>")
End sub

%>
</body>
</center>
</html>