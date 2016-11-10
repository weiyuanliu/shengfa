<% Option Explicit %>
<!--#Include file="Head.asp"-->
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
  <tr><td colspan='2' align="center"><img src="images/order_ok.jpg" width="100%"></td></tr>
<tr><td colspan="2" class="orderok_t">

<%=savedingdan()%>
<%
'货到付款
Call OK()
call sendSms(1,Trim(Request.Form("Sh_Name")),Trim(Request.Form("Sh_Tel")))
function savedingdan()
	Dim ProdName,dgtime,AddTime,Sh_Name,Sh_Mobel,Sh_Tel,Sh_Telto,Sh_ipto,Products,AddresStr,fangshi,zifu,ZipCodes,ProductNo,ipadd
    ProdName=Trim(Request.form("ProdName"))
    dgtime=Trim(request.Form("dgtime"))
	Sh_Name=Trim(Request.Form("Sh_Name"))
	'Sh_Mobel=Trim(Request.Form("Sh_Mobel"))
	Sh_Tel=Trim(Request.Form("Sh_Tel"))
	Sh_Telto=Trim(Request.Form("Sh_Telto"))
	Sh_ipto=Trim(Request.Form("Sh_ipto"))
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
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
	'
		rs("Address")=AddresStr
		rs.update
		'
		rs.close()
		set rs=Nothing
		'response.Write("<script language=javascript>"&vbcrlf)
			'response.Write("alert('不能重复提交定单！');")
			'response.Write("window.history.go(-1);")
		'response.Write("</ script>")
		'response.End()
		exit function
	end if
	rs.close()
	Dim Remark,ADS_Link
	ADS_Link = request.Cookies("advlink")
	Remark=fangshi&"|"&zifu&"|"&Trim(Request.form("SumMemony"))
	sql="insert into NwebCn_Order (ProductName,AddTime,Linkman,Address,ZipCode,Telephone,Amount,ProductNo,tel,Telto,ipto,ipaddress,Remark,ADS_Link) VALUES('"&ProdName&"','"&AddTime&"','"&Sh_Name&"','"&AddresStr&"','"&ZipCodes&"','"&Sh_Mobel&"','"&Products&"','"&ProductNo&"','"&Sh_Tel&"','"&Sh_Telto&"','"&Sh_ipto&"','"&ipadd&"','"&Remark&"',"&ADS_Link&")"
	conn.execute(sql)
	response.Cookies("advlink") = ""
end function

	Sub Ok()
		response.Write(GetValues("NwebCn_About","Content",55))
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