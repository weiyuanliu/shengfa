<%
function getipadd()
 ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
 if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR")
 getipadd=ipadd
end function
%>
       <div style="text-align:center"><img src="style/blue/images/page6_07.jpg" width="491" height="84" /></div>
       <div style="font-size:16px; font-weight:bold; color:#06F">֧�������߶���</div>
       <form action="AliPay.asp?Action=ArrayPlay" name="On_Order" id="On_Order" method="post" onsubmit="return Check_OnOrder();"> 
       <input type="hidden" name="ipadd" value="<%=getipadd%>">
       <input type="hidden"  name="On_dgtime" id="On_dgtime" value="<%=Now()%>">
       <input id="price1" value="260" size="5" name="price1" readonly="TRUE" style=" display:none">
       <input id="price2" value="360" size="5" name="price2" readonly="TRUE" style=" display:none">
		<%
			Dim OrderId,str
			Dim THISO:THISO=str&right(year(now),1)&month(now)&day(now)&XXL(5)
			OrderId = HaveOrderId(str,THISO)
			%>
       <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>">
        <table width="100%" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"  class="ordertd">
                                        <tr >
                                            <td width="18%" align="center" style=" background-color:#F0F0F0">�ջ���:&nbsp;</td>
                                            <td width="36%" style=" background-color:#F0F0F0" class="otderpa"><input class="inputone" style="width:120px;" name="On_ShName" id="On_ShName" maxlength="8"/>
                                                <font color="red">*</font>
                                            </td>
                                            <td width="46%" style=" background-color:#F0F0F0">�ֻ�:&nbsp;
                                              <input class="inputone" style="width:120px;" name="On_ShTel" id="On_ShTel" />&nbsp;<font color="red">*</font>(11λ����) </td>
                                        </tr>
                                        <tr>
                                            <td  colspan="3" align="center" nowrap="nowrap">	
                                            <font color="#FF0000">��ע�⣡��������Я�����ƶ��绰���룬�����ű��ڿ�ݹ�˾�ͻ�ʱ��ʱ����ȡ����ϵ��</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td width="18%" align="center" rowspan="2" >��Ʒ����:</td>
                                            <%Call ProdList2()%>
                                        <tr>
                                            <td width="18%" align="center">��Ʒ�ܼ�:</td>
                                            <td  colspan="2"  class="otderpa"><input id="tprice"  value="0" size="7" name="tprice" readonly="TRUE">&nbsp;Ԫ
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="center">ʡ��/����:</td>
                                            <td  colspan="2"  class="otderpa"><input  class="inputone" maxlength="10" size="10" name="On_Sheng" id="On_Sheng"/>&nbsp;ʡ&nbsp;/&nbsp;
                                              <input  class="inputone" maxlength="10" size="10" name="On_Shi" id="On_Shi"/>&nbsp;��
                                              <font color="red">*</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="center">��ϵ��ַ:</td>
                                            <td  colspan="2"  class="otderpa"><input class="inputone" maxlength="60" size="32" name="On_Addres" id="On_Addres"/>&nbsp;<font color="red">*</font>(����д��ʵ��ϵ��ַ��
                                            </td>
                                        </tr>
                                        <tr>
                                            <td  align="center">��������:</td>
                                            <td  colspan="2"  class="otderpa"><input class="inputone" maxlength="6" size="13" name="On_ZipCode" id="On_ZipCode" />
                                                <font color="red">*</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td  align="center">���ʽ:</td>
                                            <td  colspan="2"  class="otderpa">֧��������֧��<input type="hidden" name="HuiKuan" value="֧����" /></td>
                                        </tr>
                                        <tr>
                                            <td width="18%" align="center">��֤��:</td>
                                            <td  colspan="2"  class="otderpa"><input type="text" name="check_alipay" id="check_alipay" size="6" maxlength="4" class="inputone" />&nbsp;<img src="Include/newcode_alipay.asp" alt="��֤�뿴�����?����ˢ����֤��!" title="��֤�뿴�����?����ˢ����֤��!" height="22" style="cursor:pointer;margin-bottom:-6px;" onClick="this.src='Include/newcode_alipay.asp?t='+(new Date().getTime());" ></td>
                                        </tr>
                           </table>
                               <div class="orderinput"><input name="" type="image" src="style/blue/images/order_18.jpg" style=" height:31px; border:0px" /></div>
                          </form>
            <div class="orderimg"><img src="style/blue/images/order_26.jpg" width="152" height="64" /><img src="style/blue/images/order_28.jpg" width="152" height="64" /><img src="style/blue/images/order_22.jpg" width="152" height="64" /><img src="style/blue/images/order_24.jpg" width="152" height="64" /></div>
	<%if Action="" then%>
	<%if GetValues("NwebCn_About","Content",61) <> "" then%>
           <div class="oderpay">
	<%=GetValues("NwebCn_About","Content",59)%>
           </div>
	<%end if%>
	<%end if%>
      <%
  	  sub ProdList2()
    	dim rs,sql
        set rs=server.CreateObject("adodb.recordset")
		sql="select ProductName,Price,Price2,PriceText from NwebCn_Products where ViewFlag=1 order by px asc"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("���޲�Ʒ��Ϣ�����ܶ�����")
			response.Write("<input type='hidden' name='On_RecordCount' id='On_RecordCount' value='0'>")
		else
			dim i
			i=1
    		while not rs.eof 
				if i=2 then
				response.Write("<tr>")
				end if
				response.Write("<td width=""36%"" class=""otderpa""><input class=""inputone"" maxlength=""60"" size=""32"" name=""ProductName"&i&""" id=""ProductName"&i&""" value="""&rs("ProductName")&"("&rs("Price2")&"Ԫ/��)"&""" style=""width:180px;"" readonly=""readonly"" /><font color=""red"">*</font></td>")
				response.Write("<td width=""46%"">��Ʒ����:&nbsp;")
					response.Write("<select name='On_Numbers"&i&"' size='1' id='On_Numbers"&i&"' onchange=""jisuan()"">")
						response.Write("<option value='NULL' selected>��ѡ������</option>")
						response.Write("<option value='"&rs("ProductName")&"(0)'>0��</option>")
						response.Write("<option value='"&rs("ProductName")&"(1)'>1��</option>")
						response.Write("<option value='"&rs("ProductName")&"(2)'>2��</option>")
						response.Write("<option value='"&rs("ProductName")&"(3)'>3��</option>")
						response.Write("<option value='"&rs("ProductName")&"(4)'>4��</option>")
						response.Write("<option value='"&rs("ProductName")&"(5)'>5��</option>")
					response.Write("</select>")
				response.Write("<font color=""red"">*</font></td></tr>")
				rs.movenext
				i=i+1
			wend
			response.Write("<input type='hidden' name='On_RecordCount' id='On_RecordCount' value='"&rs.recordcount&"'>")
		end if
		rs.close()
		set rs=Nothing
    End sub
    %>
	
    <script language="javascript">
	<!--
	function Check_OnOrder()
	{
		var On_dgtime,On_ShName,On_ShMoble,On_ShTel,On_Sheng,On_Shi,On_Xian,On_Addres,HuiKuan,On_RecordCount,check_alipay;
		On_dgtime=document.getElementById("On_dgtime");
		On_ShName=document.getElementById("On_ShName");
		//On_ShMoble=document.getElementById("On_ShMoble");
		On_ShTel=document.getElementById("On_ShTel");
		//On_Sheng=document.getElementById("On_Sheng");
		On_Shi=document.getElementById("On_Shi");
		On_Xian=document.getElementById("On_Xian");
		On_Addres=document.getElementById("On_Addres");
		On_ZipCode=document.getElementById("On_ZipCode");
		HuiKuan=document.getElementById("HuiKuan");
		check_alipay=document.getElementById("check_alipay");
		On_RecordCount=document.getElementById("On_RecordCount");
		
		if(On_dgtime.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("���ݳ�����ˢ����ҳ��");
			return false;
		}
		if(On_ShName.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ջ�����Ϣ��");
			On_ShName.focus();
			return false;
		}
		
		/*if(On_ShMoble.value.replace(/^\s*|\s*$/g,'')!="")
		{
			var moble=On_ShMoble.value;
			var patrn1=/^[+]{0,1}(\d){1,3}[ ]?([-]?((\d)|[ ]){1,12})+$/;
			if(!patrn1.exec(moble))
			{
				alert("����д��ȷ�ĵ绰���룡");
				On_ShMoble.select();
				return false;
			}
		}
		*/
		
		if(On_ShTel.value.replace(/^\s*|\s*$/g,'')!="")
		{
		//	var tel=On_ShTel.value
		//	var patrn2=/^[+]{0,1}(\d){1,4}[ ]?([-]?((\d)|[ ]){1,12})+$/;
		//	if(!patrn2.exec(tel))
		//	{
		//		alert("����д�ջ�����ϵ��Ϣ��");
		//		On_ShTel.focus();
		//		return false;
		//	}
		}

		if(On_ShTel.value.length != "11" && On_ShTel.value[0] == "1" || On_ShTel.value.length > "13")
		{ 
			alert("����ȷ��д11λ�ֻ����룡");
			On_ShTel.focus();
			return false;
		}

		if(On_ShTel.value.length != "11")
		{ 
			alert("����ȷ��д11λ�ֻ����룡");
			On_ShTel.focus();
			return false;
		}
		if(On_Shi.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�м���Ϣ��");
			On_Shi.focus();
			return false;
		}
		if(On_Xian.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ؼ���Ϣ��");
			On_Xian.focus();
			return false;
		}
		if(On_Addres.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ջ��˵�ַ��");
			On_Addres.focus();
			return false;
		}
		if(On_ZipCode.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ʱ࣡");
			On_ZipCode.focus();
			return false;
		}
		if(check_alipay.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д��֤�룡");
			check_alipay.focus();
			return false;
		}
		if(HuiKuan.value=="NULL")
		{
			alert("����ѡ�񸶿ʽ��");
			return false;
		}
		if(parseInt(On_RecordCount.value)<=0)
		{
			alert("������Ʒ���Զ��������ܶ�����");
			return false;
		}
		else
		{
			var falg=false;
			for(var i=1;i<=parseInt(On_RecordCount.value);i++)
			{
				if(document.getElementById("On_Numbers"+i).value!="NULL")
				{
					falg=true;
				}
			}
			if(!falg)
			{
				alert("��ѡ�񶨹���Ʒ��������");
				return false;
			}
		
		}
		return true;
	}
	
	-->
	</script>