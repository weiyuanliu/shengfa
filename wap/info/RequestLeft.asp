<%

Dim dgtime,Sh_Name,Sh_Mobel,Sh_Tel,Addres,RecordCounts,ZipCodes,AddersStr,OrderId,ipadd
dgtime=trim(SafeRequest("dgtime","post"))
Sh_Name=Trim(SafeRequest("Sh_Name","post"))
Sh_Mobel=Trim(SafeRequest("Sh_Mobel","post"))
Sh_Tel=Trim(SafeRequest("Sh_Tel","post"))
Addres=Trim(SafeRequest("Addres","post"))
RecordCounts=Trim(SafeRequest("RecordCount","post"))
ZipCodes=Trim(SafeRequest("ZipCode","post"))
OrderId=SafeRequest("OrderId","post")
ipadd=SafeRequest("ipadd","post")

if dgtime="" or isnull(dgtime) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('日期出错，请返回！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if Sh_Name="" or isnull(Sh_Name) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写收货人信息！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if Sh_Tel="" or isnull(Sh_Tel) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写收货人联系信息！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if left(Sh_Tel,1)=1 and len(Sh_Tel)<>11 or len(Sh_Tel)>13 then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请正确填写11位手机号码！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if Addres="" or isnull(Addres) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请填写联系地址！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if RecordCounts<>"" then
	if Not IsNumeric(RecordCounts) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('数据出错请返回！');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>")
		response.End()
	else
		if Cint(RecordCounts)<=0 then
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

dim i,falg
falg=false
for i=1 to Cint(RecordCounts)
	if(Trim(Request.Form("Numbers"&i)))<>"NULL" then
		falg=true
	end if
next

if not falg then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('数据出错，请返回！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if
%>

<%
'电话号码
dim m_sql,m_rs,m_conn,mobileto
Set m_conn = Server.CreateObject("ADODB.Connection")
m_conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("../dbdata/Mobiledata.mdb")
m_sql="select * from Dm_Mobile where MobileNumber='"&left(Sh_Tel,7)&"'"
set m_rs = Server.CreateObject("ADODB.RecordSet")
m_rs.open m_sql,m_conn,1,1
if not m_rs.eof then
	if Instr(m_rs("MobileArea")," ") >0 then
	mobileto=m_rs("MobileArea")
	else
	mobileto=" "&m_rs("MobileArea")
	end if
end if
m_rs.Close 
set m_rs=nothing
m_conn.Close
set m_conn=nothing

'IP
u_ip=ipadd
if u_ip="" then
u_ip = Request.ServerVariables("REMOTE_ADDR") 
end if
Function cacuIp(u_ip)
On Error Resume Next
Dim srIp, aIp
srIp=0
aIp = Split(u_ip,".")
If UBound(aIP)<>3 Then
cacuIP=0
Exit Function
End If
For i=0 To 3
srIp=srIp+(CInt(aIP(i))*(256^(3-i)))
Next
cacuIp=srIp-1
If Err Then cacuIp=0
End Function 
Set iCONN=Server.CreateObject("ADODB.Connection")
iCONN.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&Server.Mappath("../dbdata/ipaddress.mdb")
iIp=cacuIp(u_ip)
iSQL = "SELECT Country FROM [IPTABLE] WHERE StartIPNum<=" & iIp & " AND EndIPNum>=" & iIp
Set rsCnt = iCONN.Execute(iSQL)
If rsCnt.Eof Then
sPlace="查无记录"
Else
sPlace=rsCnt(0)
End If
rsCnt.close()
Set rsCnt=Nothing


Function GetHttpPage(murl)
dim Http,ore,Matches
Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
Http.open "GET",murl,False
Http.Send()
If Http.Readystate<>4 and Http.status<>200 then
Set Http=Nothing
Exit function
End if
Set ore = New RegExp
ore.Pattern = "<meta[^>]+charset=[""]?([\w\-]+)[^>]*>"
ore.Global = True
ore.IgnoreCase = True
Set Matches = ore.execute(Http.responseText)
If(Matches.count>0)Then
GetHTTPPage=bytesToBSTR(Http.responseBody,Matches(0).submatches(0))
Else  
'GetHTTPPage=Http.responseText  '没有找到编码则不转换编码
GetHTTPPage=bytesToBSTR(Http.responseBody,"utf-8") '没有找到编码则转换为GB2312
End if
Set Http=Nothing
End Function

Function GetHttpPage(murlip)
dim Http,ore,Matches
Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
Http.open "GET",murlip,False
Http.Send()
If Http.Readystate<>4 and Http.status<>200 then
Set Http=Nothing
Exit function
End if
Set ore = New RegExp
ore.Pattern = "<meta[^>]+charset=[""]?([\w\-]+)[^>]*>"
ore.Global = True
ore.IgnoreCase = True
Set Matches = ore.execute(Http.responseText)
If(Matches.count>0)Then
GetHTTPPage=bytesToBSTR(Http.responseBody,Matches(0).submatches(0))
Else  
'GetHTTPPage=Http.responseText  '没有找到编码则不转换编码
GetHTTPPage=bytesToBSTR(Http.responseBody,"utf-8") '没有找到编码则转换为GB2312
End if
Set Http=Nothing
End Function

Function BytesToBstr(body,Cset)
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText
objstream.Close
set objstream = nothing
End Function

Function GetKey(HTML,Start,Last)
dim filearray,filearray2
filearray=split(HTML,Start)
if ubound(filearray)>0 then
filearray2=split(filearray(1),Last)
GetKey=filearray2(0)
end if
End Function

%>
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
<div class="order_ok" >   
<TABLE cellSpacing="1" cellpadding="1" width="90%"  align=center >
    <form name="Save_Order" id="Save_Order" method="post" action="SaveOrder.asp">
	<input type="hidden" name="Sh_ipto" value="<%=sPlace%>" />
	<input type="hidden" name="ZipCode" id="ZipCode" value="<%=ZipCodes%>" />
	<input type="hidden" name="ProductNo" id="ProductNo" value="<%=OrderId%>" />
  <tr><td colspan='2' class="orderok_t">请仔细核对订单信息</td></tr>
    <TR>
	  <TD width="30%" align=center>产品名称:</TD>
	  <TD width="70%">汉方演绎<input type="hidden" name="ProdName" value="古那迪" /></TD>
    </TR>
    <TR>
	    <TD align=center>订购时间:</TD>
	    <TD><%=FormatDate(dgtime,4)%><input type="hidden" name="dgtime" value="<%=dgtime%>" /><input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>"></TD>
    </TR>
    <TR>
    	<TD align=center>收 货 人:</TD>
    	<TD><%=Sh_Name%><input type="hidden" name="Sh_Name" value="<%=Sh_Name%>" /></TD>
    </TR>
    <TR>
	    <TD align=center>联系手机:</TD>
	    <TD><%=Sh_Tel%><input type="hidden" name="Sh_Tel" value="<%=Sh_Tel%>" /><input type="hidden" name="Sh_Telto" value="<%=mobileto%>" /></TD>
    </TR>
    <TR>
	    <TD align=center>产品清单:</TD>
	    <TD>
    	<%
		dim j,str
		for j=1 to Cint(RecordCounts)
			if Trim(Request.Form("Numbers"&j)<>"NULL") then
				response.Write(Trim(Request.Form("Numbers"&j)))
				if str="" or isnull(str) then
					str=Trim(Request.Form("Numbers"&j))
				else
					str=str&"|"&Trim(Request.Form("Numbers"&j))
				end if
				response.Write("套")
				response.Write("&nbsp;&nbsp;")			
			end if
		next
		%>
	    <input type="hidden" name="Products" value="<%=str%>" /></TD>
    </TR>
    <TR>
    	<TD align=center>收货地址:</TD>
    	<TD><input type="text" style="width:80%px;" name="AddresStr" value="<%=(AddersStr&Addres)%>" /> 地址栏内容可修改</TD>
    </TR>
    <TR>
    	<TD align=center>交易金额:</TD>
    	<TD>
    	<%=FormatCurrency(Sum_Memony(str,0),2,-2)%>元
		<input type="hidden" name="SumMemony" value="<%=Sum_Memony(str,0)%>" />
	</TD>
    </TR>
    <TR>
    	<TD align=center>付款方式:</TD>
    	<TD>货到付款<input type="hidden" name="zifu" value="货到付款" /><input type="hidden" name="fangshi" value="免费送货上门" /></TD>
    </TR> 
   <TR>
	<TD colspan="2" align=center  style=" padding-top:10px"><input type="submit" name="tijao" id="tijao" class="tijiao" value="定单提交" style="margin-right:20px;"/><input type="button" name="tijao2" id="tijao2" class="tijiao" value="返 回" onclick="window.history.go(-1);"/></TD>
  </TR>
    <TR>
    	<TD colspan='2' class="ordersm"></TD>
    </TR>
  </form>
</table>
</div>
<%
dim ProductNo
ProductNo=OrderId
%>
<%=savedingdan()%>
<%

function savedingdan()

	dim rs,sql
	'ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		'if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR") 
	 set rs=server.CreateObject("adodb.recordset")
	sql="select * from NwebCn_Order where ProductNo='"&ProductNo&"'"
	rs.open sql,conn,1,1
	if not rs.eof and not rs.bof then
		rs.close()
		set rs=Nothing
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('不能重复提交定单！');")
			response.Write("window.history.go(-1);")
			Response.Write("</script>") 
	 	response.End()
	 	exit function
	 end if
	 rs.close()
	 
	sql="select top 1 * from NwebCn_Order"
	rs.open sql,conn,1,3
	rs.addnew()
	rs("ProductName")="古那迪"
	rs("AddTime")=dgtime
	rs("Linkman")=Sh_Name
	rs("Address")= AddersStr&Addres
	rs("ZipCode")=ZipCodes
	rs("Amount")=str
	rs("ProductNo")=ProductNo
	rs("HuoDao_FuKuan")=true
	rs("tel")=Sh_Tel
	rs("telto")=mobileto
	rs("ipaddress")=ipadd
	rs("ipto")=sPlace
	rs("ADS_Link") = request.Cookies("advlink")
	rs("Remark")="免费送货上门"&"|货到付款|"&Sum_Memony(str,0)

	'查询是否在黑名单内
	dim rsbl,sqlbl
	set rsbl=server.CreateObject("adodb.recordset")
	sqlbl="select * from NwebCn_Order where Tel='"&Sh_Tel&"' and Fax='0' and blacklist='1'"
	rsbl.open sqlbl,conn,1,1
	if not rsbl.eof and not rsbl.bof then
	rs("blacklist")="2"
	end if
	rsbl.close()

	rs.update()
	rs.close()
	set rs=Nothing
	'Call OK()
end function

function ding_No()
	dim shijian
  shijian=now()
   ding_NO=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
	'ding_NO=FormatDate(now,3)
end function

Sub Ok()
	response.Write("<div style='width:98%; height:300px; background:#ffffff; padding:10px; border:#000000 1px solid;'>")
		response.Write("<div style='text-align:center; font-size:18px; color:#0351A3; padding-top:10px; padding-bottom:10px;'><strong>订单已经提交成功！！</strong></div>")
		response.Write("<div style='text-align:left; line-height:25px;'>")
			response.Write(GetValues("NwebCn_About","Content",55))
		response.Write("</div>")
		response.Write("<div style='padding-top:10px;'>")
			response.Write("<input type='button' name='getbak' id='getbak' value='返 回' style='border:#0351A3 1px solid; font-size:14px; padding-top:3px;' onclick=""window.location.href='index.asp';"">")
		response.Write("</div>")
	response.Write("</div>")
End sub
 


	Function Sum_Memony(str,str1)
		Dim Str2,Prodname,Numbers,Sum,i
		Dim rs,sql
		Set rs=server.CreateObject("adodb.recordset")
		Str2=Split(str,"|")
		Sum=0
		for i=0 to ubound(Str2)
			sql="select top 1 Price from NwebCn_Products where ProductName='"&Mid(str2(i),1,instr(str2(i),"(")-1)&"'"
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
				Sum=Sum+(rs("Price")*Cint(Mid(str2(i),instr(Str2(i),"(")+1,1)))
			end if
			rs.close()
		next
		Sum=Sum+Str1
		Sum_Memony=Sum
	End Function
%>