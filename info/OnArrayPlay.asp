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
	On_ipadd=SafeRequest("ipadd","post")

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

if left(On_ShTel,1)=1 and len(On_ShTel)<>11 or len(On_ShTel)>13 then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('请正确填写11位手机号码！');"&vbcrlf)
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
	On_AdderStr=On_Sheng&"省"&On_Shi&"市"
else
	On_AdderStr=On_Shi&"市"
end if

if session("firstecode_alipay")<>request("check_alipay") then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('您输入验证码错误，请返回重新输入！');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if
%>

<%
'电话号码
dim m_sql,m_rs,m_conn,mobileto
Set m_conn = Server.CreateObject("ADODB.Connection")
m_conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("dbdata/Mobiledata.mdb")
m_sql="select * from Dm_Mobile where MobileNumber='"&left(On_ShTel,7)&"'"
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
u_ip=On_ipadd
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
iCONN.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&Server.Mappath("dbdata/ipaddress.mdb")
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

Function GetHttpPage(murla)
dim Http,ore,Matches
Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
Http.open "GET",murla,False
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

'tel
'dim murl,StartGet,Teltoprovince,Teltocity,Sh_Telto
'murl="http://www.baidu.com/s?wd="&On_ShTel
'StartGet = getHTTPPage(murl)
'Teltoprovince=Getkey(StartGet,On_ShTel&"&quot;</span>                <span>","</span>")
'if Teltoprovince <> "" then
'if Instr(Teltoprovince,"&nbsp;") >0 then
'	Sh_Telto=replace(Teltoprovince,"&nbsp;"," ")
'else
'	Sh_Telto=Teltoprovince
'end if
'end if

'IP
'dim murla,StartGeta,On_iptoadd,On_ipto
'murla="http://www.baidu.com/s?wd="&On_ipadd
'StartGeta = getHTTPPage(murla)
'On_iptoadd=Getkey(StartGeta,"IP地址:&nbsp;"&On_ipadd&"</span>","</td>")
'if On_iptoadd <> "" then
'	On_ipto=On_iptoadd
'end if
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
<TABLE  cellSpacing="1" cellpadding="1" width="90%"  align=center >
<form name="Save_Order" id="Save_Order" method="post" action="SaveOrder2.asp">
<input type="hidden" name="On_ipto" value="<%=sPlace%>" />
  <tr><td colspan='2' class="orderok_t">请仔细核对订单信息</td></tr>
    <TR>
	  <TD width="30%" align=center>产品名称:</TD>
	  <TD width="70%">倍洛加<input type="hidden" name="ProdName" value="倍洛加" /></TD>
    </TR>
    <TR>
	    <TD align=center>订购时间:</TD>
	    <TD><%=FormatDate(On_dgtime,4)%><input type="hidden" name="dgtime" value="<%=On_dgtime%>" /><input type="hidden" name="OrderId" value="<%=OrderId%>" /></TD>
    </TR>
    <TR>
    	<TD align=center>收 货 人:</TD>
    	<TD><%=On_ShName%><input type="hidden" name="Sh_Name" value="<%=On_ShName%>" /></TD>
    </TR>
    <TR>
	    <TD align=center>联系手机:</TD>
	    <TD><%=On_ShTel%><input type="hidden" name="Sh_Tel" value="<%=On_ShTel%>" /><input type="hidden" name="Sh_Telto" value="<%=mobileto%>" /></TD>
    </TR>
    <TR>
	    <TD align=center>产品清单:</TD>
	    <TD>
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
				response.Write("&nbsp;&nbsp;")			
			end if
		next
		%>
        <input type="hidden" name="Products" value="<%=str2%>" /></TD>
    </TR>
    <TR>
    	<TD align=center>邮政编码:</TD>
    	<TD><%=On_ZipCode%><input type="hidden" name="ZipCode" id="ZipCode" value="<%=On_ZipCode%>" /></TD>
    </TR>
    <TR>
    	<TD align=center>收货地址:</TD>
    	<TD><input type="text" style="width:300px;" name="AddresStr" value="<%=(On_AdderStr&On_Addres)%>" /> 地址栏内容可修改</TD>
    </TR>
    <TR>
    	<TD align=center>交易金额:</TD>
    	<TD>
	<%=FormatCurrency(Sum_Memony2(str2,0),2,-2)%>元<input type="hidden" name="SumMemony" value="<%=Sum_Memony2(str2,0)%>" />
	</TD>
    </TR>
    <TR>
    	<TD align=center>付款方式:</TD>
    	<TD><%=Trim(Request.Form("HuiKuan"))%><input type="hidden" name="zifu" value="<%=Trim(Request.Form("HuiKuan"))%>" /></TD>
    </TR>  
<input type="hidden" name="ProductNo" id="ProductNo" value="<%=OrderId%>" />
   <TR>
	<TD colspan="2" align=center  style=" padding-top:10px"><input type="submit" name="tijao" id="tijao" class="tijiao" value="定单提交" style="margin-right:20px;"/><input type="button" name="tijao2" id="tijao2" class="tijiao" value="返 回" onclick="window.history.go(-1);"/></TD>
  </TR>
    <TR>
    	<TD colspan='2' class="ordersm"></TD>
    </TR>
</form>
</TABLE>
</div>

<%
DIm productno1
productno1 =OrderId

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
	rs("ProductNo")=OrderId
	rs("tel")=On_ShTel
	rs("telto")=mobileto
	rs("ipaddress")=On_ipadd
	rs("ipto")=sPlace

	'if instr(zifu,"支付宝")>0 or instr(zifu,"网银")>0 then
		'rs("State")="已付款"
	'else
		'rs("State")="未付款"
	'end if
	rs("Remark")="免费快递送货上门|"&Trim(Request.Form("HuiKuan")&"|"&Sum_Memony2(str2,0))

	'查询是否在黑名单内
	dim rsbl,sqlbl
	set rsbl=server.CreateObject("adodb.recordset")
	sqlbl="select * from NwebCn_Order where Tel='"&On_ShTel&"' and Fax='0' and blacklist='1'"
	rsbl.open sqlbl,conn,1,1
	if not rsbl.eof and not rsbl.bof then
	rs("blacklist")="2"
	end if
	rsbl.close()

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