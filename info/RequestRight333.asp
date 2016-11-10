<%
Dim On_dgtime,On_ShName,On_ShMoble,On_ShTel,On_Sheng,On_Shi,On_Xian,On_ZipCode,On_Addres,HuiKuan,On_RecordCount,On_XOrStr,On_QuType,On_ipadd
Dim On_AdderStr

	On_dgtime=Trim(SafeRequest("On_dgtime","post"))
	On_ShName=Trim(SafeRequest("On_ShName","post"))
	On_ShMoble=Trim(SafeRequest("On_ShMoble","post"))
	On_ShTel=Trim(SafeRequest("On_ShTel","post"))
	On_Sheng=trim(SafeRequest("On_Sheng","post"))
	On_Shi=Trim(SafeRequest("On_Shi","post"))
	On_Xian=Trim(SafeRequest("On_Xian","post"))
	On_ZipCode=Trim(SafeRequest("On_ZipCode","post"))
	On_Addres=Trim(SafeRequest("On_Addres","post"))
	HuiKuan=Trim(SafeRequest("HuiKuan","post"))
	On_RecordCount=Trim(SafeRequest("On_RecordCount","post"))
	On_QuType=Trim(SafeRequest("On_QuType","post"))
	OrderId=SafeRequest("OrderId","post")
	On_ipadd=SafeRequest("On_ipadd","post")

if On_QuType="1" then
	On_XOrStr="区"
else
	On_XOrStr="县"
end if

if On_dgtime="" or isnull(On_dgtime) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('日期出错，请返回！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if On_ShName="" or isnull(On_ShName) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写收货人信息！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if On_ShTel="" or isnull(On_ShTel) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写收货人联系信息！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if On_Shi="" or isnull(On_Shi) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写市级信息！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if On_Xian="" or isnull(On_Xian) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写县级信息！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if On_Addres="" or isnull(On_Addres) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写联系地址！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if On_RecordCount<>"" then
	if Not IsNumeric(On_RecordCount) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('数据出错请返回！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>")
		response.End()
	else
		if Cint(On_RecordCount)<=0 then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('数据出错请返回！');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>")
			response.End()
		end if
	end if
else
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('数据出错请返回！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

dim ii,falg2
falg2=false
for ii=1 to Cint(On_RecordCount)
	if(Trim(Request.Form("On_Numbers"&ii)))<>"NULL" then
		falg2=true
	end if
next

if not falg2 then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('数据出错，请返回！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if On_Sheng <> "" then
	On_AdderStr=On_Sheng&"省"&On_Shi&"市"&On_Xian
else
	On_AdderStr=On_Shi&"市"&On_Xian
end if
%><style type="text/css">
<!--
.STYLE1 {color: #0351A3}
.STYLE3 {color: #0351A3; font-weight: bold; }
-->
</style>
<div class="Order_Text" style="border:#0351A3 1px solid;background:#E1F2FA;margin:15px; width:60%; text-align:left;">
	<table width="97%" border="0" align="center" cellpadding="5" cellspacing="0" style="margin-top:10px; margin-bottom:10px;">
    <form name="Save_Order" id="Save_Order" method="post" action="SaveOrder2.asp">
  <tr>
    <td height="32" colspan="2" align="center" style="border-bottom:#0351A3 1px solid;"><span class="STYLE3">请 仔 细 核 对 订 单 信 息</span></td>
    </tr>
  <tr>
    <td width="22%" height="30" align="right"><span class="STYLE1">产吕名称：</span></td>
    <td width="78%" height="30"><span class="STYLE1">倍洛加<input type="hidden" name="ProdName" value="倍洛加" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">订购时间：</span></td>
    <td height="30"><span class="STYLE1 STYLE1"><%=FormatDate(On_dgtime,4)%><input type="hidden" name="dgtime" value="<%=On_dgtime%>" />
    <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>"></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">收货人信息：</span></td>
    <td height="30"><span class="STYLE1 STYLE1"><%=On_ShName%><input type="hidden" name="Sh_Name" value="<%=On_ShName%>" /></span></td>
  </tr>
  
  <tr>
    <td height="30" align="right"><span class="STYLE1">联系方式：</span></td>
    <td height="30"><span class="STYLE1"><%=On_ShTel%><input type="hidden" name="Sh_Tel" value="<%=On_ShTel%>" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">订购商品：</span></td>
    <td height="30"><span class="STYLE1"></span></td>
  </tr>
  <tr>
    <td height="30" align="right">&nbsp;</td>
    <td height="30" style="line-height:20px;">
    	<%
		dim jj,str2
		for jj=1 to Cint(On_RecordCount)
			if Trim(Request.Form("On_Numbers"&jj)<>"NULL") then
				response.Write(Trim(Request.Form("On_Numbers"&jj)))
				if str2="" or isnull(str2) then
					str2=Trim(Request.Form("On_Numbers"&jj))
				else
					str2=str2&"|"&Trim(Request.Form("On_Numbers"&jj))
				end if
				response.Write("盒")
				response.Write("<br/>")			
			end if
		next
		%>
        <input type="hidden" name="Products" value="<%=str2%>" />    </td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">收货地址：</span></td>
    <td height="30"><span class="STYLE1"><%'=On_AdderStr%><%'=On_XOrStr%><%'=On_Addres%><input type="text" style="width:300px;" name="AddresStr" value="<%=(On_AdderStr&On_XOrStr&On_Addres)%>" /> 地址栏内容可修改</span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">邮政编码：</span></td>
    <td height="30"><span class="STYLE1"><%=On_ZipCode%><input type="hidden" name="ZipCode" id="ZipCode" value="<%=On_ZipCode%>" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">送货方式：</span></td>
    <td height="30"><span class="STYLE1">免费快递送货上门<input type="hidden" name="fangshi" value="免费快递送货上门" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">支付方式：</span></td>
    <td height="30"><span class="STYLE1"><%=Trim(Request.Form("HuiKuan"))%></span><input type="hidden" name="zifu" value="<%=Trim(Request.Form("HuiKuan"))%>" /></td>
  </tr>
   <tr>
    <td height="30" align="right"><span class="STYLE1">总 金 额：</span></td>
    <td height="30">
    	<%=FormatCurrency(Sum_Memony2(str2,0),2,-2)%>元
		<input type="hidden" name="SumMemony" value="<%=Sum_Memony2(str2,0)%>" />	</td>
  </tr>
  
  <tr>
    <td height="30" colspan="2" align="center"><label>
    	<input type="hidden" name="ProductNo" id="ProductNo" value="<%=OrderId%>" />
      <input type="submit" name="tijao" id="tijao" value="定单提交" style="margin-right:20px;font-size:14px; padding-top:3px; border:#0351A3 1px solid;"/>
    </label>      <input type="button" name="tijao2" id="tijao2" value="返 回" style="margin-right:20px; font-size:14px; padding-top:3px;border:#0351A3 1px solid;" onclick="window.history.go(-1);"/></td>
    </tr>
  </form>
</table>
</div>
<%
DIm productno1
productno1 =ding_No()
%>
<%'=savedingdan1()%>
<%
 
function savedingdan1()
	 
	
	dim rs,sql
	'ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		'if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR") 
	
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from NwebCn_Order where ProductNo='"&ProductNo1&"'"
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
	rs("ProductName")="倍洛加"
	rs("AddTime")=now()
	rs("Linkman")=On_ShName
	rs("Address")=On_AdderStr&On_XOrStr&On_Addres
	rs("ZipCode")=On_ZipCode
	'rs("Telephone")=Sh_Mobel
	rs("Amount")=str2
	rs("ProductNo")=ProductNo1
	rs("tel")=On_ShTel

	rs("ipaddress")=On_ipadd

	'if instr(zifu,"支付宝")>0 or instr(zifu,"网银")>0 then
		'rs("State")="已付款"
	'else
		'rs("State")="未付款"
	'end if
	rs("ADS_Link") = request.Cookies("advlink")
	rs("Remark")="免费快递送货上门|"&Trim(Request.Form("HuiKuan")&"|"&Sum_Memony2(str2,0))
	rs.update()
	rs.close()
	set rs=Nothing
	end function
 
	Function Sum_Memony2(str,str1)
		Dim Str2,Prodname,Numbers,Sum,i
		Dim rs,sql
		Set rs=server.CreateObject("adodb.recordset")
		Str2=Split(str,"|")
		Sum=0
		for i=0 to ubound(Str2)
			sql="select top 1 Price2 from NwebCn_Products where ProductName='"&Mid(str2(i),1,instr(str2(i),"(")-1)&"'"
			rs.open sql,conn,1,1
			if rs.eof and rs.bof then
				rs.close()
				set rs=Nothing
				response.Write("<script languge=javascript>")
					response.Write("alert('数据出错，请返回！');")
					response.Write("window.history.go(-1);")
				response.Write("</script>")
				response.End()
				exit function
			else
				Sum=Sum+(rs("Price2")*Cint(Mid(str2(i),instr(Str2(i),"(")+1,1)))
			end if
			rs.close()
		next
		Sum=Sum+Str1
		Sum_Memony2=Sum
	End Function
%>