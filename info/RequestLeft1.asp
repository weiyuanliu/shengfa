<%
Dim dgtime,Sh_Name,Sh_Mobel,Sh_Tel,Sheng,shi,xian,Addres,RecordCounts,ZipCodes,QuType,XOrStr,AddersStr,OrderId,ipadd
dgtime=trim(SafeRequest("dgtime","post"))
Sh_Name=Trim(SafeRequest("Sh_Name","post"))
Sh_Mobel=Trim(SafeRequest("Sh_Mobel","post"))
Sh_Tel=Trim(SafeRequest("Sh_Tel","post"))
Sheng=Trim(SafeRequest("Sheng","post"))
shi=Trim(SafeRequest("shi","post"))
xian=Trim(SafeRequest("xian","post"))
Addres=Trim(SafeRequest("Addres","post"))
RecordCounts=Trim(SafeRequest("RecordCount","post"))
ZipCodes=Trim(SafeRequest("ZipCode","post"))
QuType=Trim(SafeRequest("QuType","post"))
OrderId=SafeRequest("OrderId","post")
ipadd=SafeRequest("ipadd","post")

if QuType="1" then
	XOrStr="��"
else
	XOrStr="��"
end if


if dgtime="" or isnull(dgtime) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('���ڳ����뷵�أ�');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if Sh_Name="" or isnull(Sh_Name) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('����д�ջ�����Ϣ��');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if Sh_Tel="" or isnull(Sh_Tel) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('����д�ջ�����ϵ��Ϣ��');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if shi="" or isnull(shi) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('����д�м���Ϣ��');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if xian="" or isnull(xian) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('����д�ؼ���Ϣ��');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if Addres="" or isnull(Addres) then
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('����д��ϵ��ַ��');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if RecordCounts<>"" then
	if Not IsNumeric(RecordCounts) then
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('���ݳ����뷵�أ�');"&vbcrlf)
			response.Write("window.history.go(-1);"&vbcrlf)
		response.Write("</script>")
		response.End()
	else
		if Cint(RecordCounts)<=0 then
			response.Write("<script language=javascript>"&vbcrlf)
				response.Write("alert('���ݳ����뷵�أ�');"&vbcrlf)
				response.Write("window.history.go(-1);"&vbcrlf)
			response.Write("</script>")
			response.End()
		end if
	end if
else
	response.Write("<script language=javascript>"&vbcrlf)
		response.Write("alert('���ݳ����뷵�أ�');"&vbcrlf)
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
		response.Write("alert('���ݳ����뷵�أ�');"&vbcrlf)
		response.Write("window.history.go(-1);"&vbcrlf)
	response.Write("</script>")
	response.End()
end if

if Sheng <> "" then
	AddersStr=Sheng&"ʡ"&shi&"��"&xian
else
	AddersStr=shi&"��"&xian
end if
%><style type="text/css">
<!--
.STYLE1 {color: #0250A2}
.STYLE3 {color: #0250A2; font-weight: bold; }
-->
</style>
<div class="Order_Text" style="border:#0351A3 1px solid;background:#EEF7FE;margin:15px; width:60%; text-align:left;">
	<table width="97%" border="0" align="center" cellpadding="5" cellspacing="0" style="margin-top:10px; margin-bottom:10px;">
    <form name="Save_Order" id="Save_Order" method="post" action="SaveOrder.asp">
  <tr>
    <td height="32" colspan="2" align="center" style="border-bottom:#4D4D4D 1px solid;"><span class="STYLE3">�� �� ϸ �� �� �� �� �� Ϣ</span></td>
    </tr>
  <tr>
    <td width="21%" height="30" align="right"><span class="STYLE1">��Ʒ���ƣ�</span></td>
    <td width="79%" height="30"><span class="STYLE1">���ǵ�
      <input type="hidden" name="ProdName" value="���ǵ�" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">����ʱ�䣺</span></td>
    <td height="30"><span class="STYLE1 STYLE1"><%=FormatDate(dgtime,4)%><input type="hidden" name="dgtime" value="<%=dgtime%>" /></span>
              <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>">
              </td> 
    </td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">�ջ�����Ϣ��</span></td>
    <td height="30"><span class="STYLE1 STYLE1"><%=Sh_Name%><input type="hidden" name="Sh_Name" value="<%=Sh_Name%>" /></span></td>
  </tr>
  <!--
  <tr>
    <td height="30" align="right"><span class="STYLE1">�ջ�����ϵ�ֻ���</span></td>
    <td height="30"><span class="STYLE1 STYLE1"><%=Sh_Mobel%><input type="hidden" name="Sh_Mobel" value="<%=Sh_Mobel%>" /></span></td>
  </tr>
  -->
  <tr>
    <td height="30" align="right"><span class="STYLE1">��ϵ��ʽ��</span></td>
    <td height="30"><span class="STYLE1"><%=Sh_Tel%><input type="hidden" name="Sh_Tel" value="<%=Sh_Tel%>" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">������Ʒ��</span></td>
    <td height="30"><span class="STYLE1"></span></td>
  </tr>
  <tr>
    <td height="30" align="right">&nbsp;</td>
    <td height="30" style="line-height:20px;">
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
				response.Write("��")
				response.Write("<br/>")			
			end if
		next
		%>
        <input type="hidden" name="Products" value="<%=str%>" />    </td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">�ջ���ַ��</span></td>
    <td height="30"><span class="STYLE1"><%'=AddersStr%><%'=XOrStr%><%'=Addres%><input type="text" style="width:300px;" name="AddresStr" value="<%=(AddersStr&XOrStr&Addres)%>" /> ��ַ�����ݿ��޸�</span></td>
  </tr>
  <tr>
    <td height="30" align="right" width="120"><span class="STYLE1">�������룺</span></td>
    <td height="30"><span class="STYLE1"><%=ZipCodes%><input type="hidden" name="ZipCode" id="ZipCode" value="<%=ZipCodes%>" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">�ͻ���ʽ��</span></td>
    <td height="30"><span class="STYLE1">����ͻ�����<input type="hidden" name="fangshi" value="����ͻ�����" /></span></td>
  </tr>
  <tr>
    <td height="30" align="right"><span class="STYLE1">֧����ʽ��</span></td>
    <td height="30"><span class="STYLE1">�������� (���л�����Ż�)</span><input type="hidden" name="zifu" value="��������" /></td>
  </tr>
   <tr>
    <td height="30" align="right"><span class="STYLE1">�� �� �</span></td>
    <td height="30">
    	<%=FormatCurrency(Sum_Memony(str,0),2,-2)%>Ԫ
		<input type="hidden" name="SumMemony" value="<%=Sum_Memony(str,0)%>" />
	</td>
  </tr>
  
  <tr>
    <td height="30" colspan="2" align="center"><label>
   	  <input type="hidden" name="ProductNo" id="ProductNo" value="<%=OrderId%>" />
      <input type="submit" name="tijao" id="tijao" value="�����ύ" style="margin-right:20px; font-size:14px; border:#0351A3 1px solid; padding-top:3px;"/>
    </label>    <input type="button" name="tijao2" id="tijao2" value="�� ��" style="margin-right:20px; font-size:14px; border:#0351A3 1px solid; padding-top:3px;" onclick="window.history.go(-1);"/></td>
  </tr>
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
			response.Write("alert('�����ظ��ύ������');")
			response.Write("window.history.go(-1);")
			Response.Write("</script>") 
	 	response.End()
	 	exit function
	 end if
	 rs.close()
	 
	sql="select top 1 * from NwebCn_Order"
	rs.open sql,conn,1,3
	rs.addnew()
	rs("ProductName")="���ǵ�"
	rs("AddTime")=dgtime
	rs("Linkman")=Sh_Name
	rs("Address")= AddersStr&XOrStr&Addres
	rs("ZipCode")=ZipCodes
	'rs("Telephone")=Sh_Mobel
	rs("Amount")=str
	rs("ProductNo")=ProductNo
	rs("HuoDao_FuKuan")=true
	rs("tel")=Sh_Tel
	rs("ipaddress")=ipadd
	rs("ADS_Link") = request.Cookies("advlink")
	rs("Remark")="����ͻ�����"&"|��������|"&Sum_Memony(str,0)
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
	response.Write("<div style='width:400px; height:300px; background:#ffffff; padding:10px; border:#000000 1px solid;'>")
		response.Write("<div style='text-align:center; font-size:18px; color:#0351A3; padding-top:10px; padding-bottom:10px;'><strong>�����Ѿ��ύ�ɹ�����</strong></div>")
		response.Write("<div style='text-align:left; line-height:25px;'>")
			response.Write(GetValues("NwebCn_About","Content",55))
		response.Write("</div>")
		response.Write("<div style='padding-top:10px;'>")
			response.Write("<input type='button' name='getbak' id='getbak' value='�� ��' style='border:#0351A3 1px solid; font-size:14px; padding-top:3px;' onclick=""window.location.href='index.asp';"">")
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
					response.Write("alert('���ݳ����뷵�أ�');")
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