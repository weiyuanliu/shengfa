<%
function getipadd()
 ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
 if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR")
 getipadd=ipadd
end function
%>
       <div style="text-align:center"><img src="style/blue/images/page6_07.jpg" width="491" height="84" /></div>
       <div style="font-size:16px; font-weight:bold; color:#06F">�����ύ����</div>
       <form action="order.asp?Action=Left" name="order1" id="order1" method="post" onsubmit="return CheckOrder1();"> 
       <input type="hidden" name="ipadd" value="<%=getipadd%>">
       <input type="hidden" name="dgtime" value="<%=Now()%>">
       <input id="price1" value="280" size="5" name="price1" readonly="TRUE" style=" display:none">
       <input id="price2" value="380" size="5" name="price2" readonly="TRUE" style=" display:none">
		<%
		Dim THISO:THISO=str&right(year(now),1)&month(now)&day(now)&XXL(5)
		OrderId = HaveOrderId(str,THISO)
		%>
       <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>">
        <table width="100%" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"  class="ordertd">
                                        <tr >
                                            <td width="18%" align="center" style=" background-color:#F0F0F0">�ջ���:&nbsp;</td>
                                            <td width="36%" style=" background-color:#F0F0F0" class="otderpa"><input class="inputone" style="width:120px;" name="Sh_Name" id="Sh_Name" maxlength="8" />
                                                <font color="red">*</font>
                                            </td>
                                            <td width="46%" style=" background-color:#F0F0F0">�ֻ�:&nbsp;
                                              <input class="inputone" style="width:120px;" name="Sh_Tel" id="Sh_Tel" />&nbsp;<font color="red">*</font>(11λ����) </td>
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
                                        </tr>                                        <tr>
                                            <td align="center">ʡ��/����:</td>
                                            <td  colspan="2"  class="otderpa"><input  class="inputone" maxlength="10" size="10" name="Sheng" id="Sheng" />&nbsp;ʡ&nbsp;/&nbsp;
                                              <input  class="inputone" maxlength="10" size="10" name="shi" id="shi" />&nbsp;��
                                              <font color="red">*</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="center">��ϵ��ַ:</td>
                                            <td  colspan="2"  class="otderpa"><input class="inputone" maxlength="60" size="32" name="Addres" id="Addres" />&nbsp;<font color="red">*</font>(����д��ʵ��ϵ��ַ��
                                            </td>
                                        </tr>
                                        <tr>
                                            <td  align="center">��������:</td>
                                            <td  colspan="2"  class="otderpa"><input class="inputone" maxlength="6" size="13" name="ZipCode" id="ZipCode" />
                                                <font color="red">*</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td  align="center">���ʽ:</td>
                                            <td  colspan="2"  class="otderpa">
                                                    <select class="inputone" size="1" name="HuiKuan" id="HuiKuan" onchange="jisuanpay()">
                                                        <option value="��������" selected="selected">��������</option>
                                                    </select> &nbsp;<strong>
                                        </tr>
                                        <tr>
                                            <td width="18%" align="center">��֤��:</td>
                                            <td  colspan="2"  class="otderpa"><input type="text" name="check_left" id="check_left" size="6" maxlength="4" class="inputone" />&nbsp;<img src="Include/newcode_left.asp" alt="��֤�뿴�����?����ˢ����֤��!" title="��֤�뿴�����?����ˢ����֤��!" height="22" style="cursor:pointer;margin-bottom:-6px;" onClick="this.src='Include/newcode_left.asp?t='+(new Date().getTime());" ></td>
                                        </tr>
                           </table>
                               <div class="orderinput"><input name="" type="image" src="style/blue/images/order_18.jpg" style=" height:31px; border:0px" /></div>
                          </form>
            <div class="orderimg"><img src="style/blue/images/order_26.jpg" width="152" height="64" /><img src="style/blue/images/order_28.jpg" width="152" height="64" /><img src="style/blue/images/order_22.jpg" width="152" height="64" /><img src="style/blue/images/order_24.jpg" width="152" height="64" /></div>
	<%if Action="" then%>
	<%if GetValues("NwebCn_About","Content",61) <> "" then%>
           <div class="oderpay">
	<%=GetValues("NwebCn_About","Content",57)%>
	<%=GetValues("NwebCn_About","Content",61)%>
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
				response.Write("<td width=""36%"" class=""otderpa""><input class=""inputone"" maxlength=""60"" size=""32"" name=""ProductName"&i&""" id=""ProductName"&i&""" value="""&rs("ProductName")&"("&rs("Price")&"Ԫ/��)"&""" style=""width:180px;"" readonly=""readonly"" /><font color=""red"">*</font></td>")
				response.Write("<td width=""46%"">��Ʒ����:&nbsp;")
					response.Write("<select name='Numbers"&i&"' size='1' id='Numbers"&i&"' onchange=""jisuan()"">")
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
			response.Write("<input type='hidden' name='RecordCount' id='RecordCount' value='"&rs.recordcount&"'>")
		end if
		rs.close()
		set rs=Nothing
    End sub
    %>
      
<%
    sub ProdList()
    	dim rs,sql
        set rs=server.CreateObject("adodb.recordset")
		sql="select ProductName,Price,Price2,PriceText from NwebCn_Products where ViewFlag=1 order by px asc"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("���޲�Ʒ��Ϣ�����ܶ�����")
			response.Write("<input type='hidden' name='RecordCount' id='RecordCount' value='0'>")
		else
			dim i
			i=1
    		while not rs.eof 
				response.Write(rs("ProductName")&"��")
				response.Write("<label>")
					response.Write("<select name='Numbers"&i&"' size='1' id='Numbers"&i&"'>")
						
						response.Write("<option value='NULL' selected>��ѡ������</option>")
						response.Write("<option value='"&rs("ProductName")&"(0)'>���</option>")
						response.Write("<option value='"&rs("ProductName")&"(1)'>һ��</option>")
						response.Write("<option value='"&rs("ProductName")&"(2)'>����</option>")
						response.Write("<option value='"&rs("ProductName")&"(3)'>����</option>")
						response.Write("<option value='"&rs("ProductName")&"(4)'>�ĺ�</option>")
						response.Write("<option value='"&rs("ProductName")&"(5)'>���</option>")
					response.Write("</select>")
				response.Write("</label>")
				response.Write("��<font color='#ff0000'>"&rs("Price")&rs("PriceText")&"</font><font color='#ff0000'>��ѿ���ͻ�</font>��")
				response.Write("<br />")
				rs.movenext
				i=i+1
			wend
			response.Write("<input type='hidden' name='RecordCount' id='RecordCount' value='"&rs.recordcount&"'>")
		end if
		rs.close()
		set rs=Nothing
    End sub
    %>
	<script language="javascript">
		function chekNum(obj){
			obj.value = obj.value.replace(/[^\d.]/g,"");
			obj.value = obj.value.replace(/^\./g,"");
			obj.value = obj.value.replace(/\.{2,}/g,".");
			obj.value = obj.value.replace(".","$#$").replace(/\./g,"").replace("$#$",".");
			if(obj.value.length > 11){
                	obj.value = obj.value.substring(0,11);      
			}
		}
		function CheckOrder1()
		{
			var Sh_Name,Sh_Tel,Sheng,shi,Addres,RecordCount,ZipCode
			Sh_Name=document.getElementById("Sh_Name");
			ShTel=document.getElementById("Sh_Tel");
			shi=document.getElementById("shi");
			Addres=document.getElementById("Addres");
			ZipCode=document.getElementById("ZipCode");
			RecordCount=document.getElementById("RecordCount");
			check_left=document.getElementById("check_left");
	
		if(Sh_Name.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ջ�����Ϣ��");
			Sh_Name.focus();
			return false;
		}
		if(ShTel.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ֻ����룡");
			ShTel.focus();
			return false;
		}
		if(ShTel.value.length != "11" && ShTel.value[0] == "1" || ShTel.value.length > "13")
		{ 
			alert("����ȷ��д11λ�ֻ����룡");
			ShTel.focus();
			return false;
		}
		if(shi.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�м���Ϣ��");
			shi.focus();
			return false;
		}
		if(Addres.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ջ��˵�ַ��Ϣ��");
			Addres.focus();
			return false;
		}
		if(ZipCode.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ʱ࣡");
			ZipCode.focus();
			return false;
		}
		if(check_left.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("������֤�룡");
			check_left.focus();
			return false;
		}
		if(parseInt(RecordCount.value)<=0)
		{
			alert("������Ʒ��Ϣ���޷�������")
			return false;
		}
		else
		{
			var falg=false;
			for(var i=1;i<=parseInt(RecordCount.value);i++)
			{
				if(document.getElementById("Numbers"+i).value!="NULL")
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
	</script>