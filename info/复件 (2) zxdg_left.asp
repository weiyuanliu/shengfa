<%
function getipadd()
 ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
 if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR")
 getipadd=ipadd
end function
%>
       <div style="text-align:center"><img src="style/blue/images/page6_07.jpg" width="491" height="84" /></div>
       <div style="font-size:16px; font-weight:bold; color:#06F">在线提交订单</div>
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
                                            <td width="18%" align="center" style=" background-color:#F0F0F0">收货人:&nbsp;</td>
                                            <td width="36%" style=" background-color:#F0F0F0" class="otderpa"><input class="inputone" style="width:120px;" name="Sh_Name" id="Sh_Name" maxlength="8" />
                                                <font color="red">*</font>
                                            </td>
                                            <td width="46%" style=" background-color:#F0F0F0">手机:&nbsp;
                                              <input class="inputone" style="width:120px;" name="Sh_Tel" id="Sh_Tel" />&nbsp;<font color="red">*</font>(11位数字) </td>
                                        </tr>
                                        <tr>
                                            <td  colspan="3" align="center" nowrap="nowrap">	
                                            <font color="#FF0000">（注意！请留随身携带的移动电话号码，这样才便于快递公司送货时及时与您取得联系）</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td width="18%" align="center" rowspan="2" >商品类型:</td>
                                            <%Call ProdList2()%>
                                        <tr>
                                            <td width="18%" align="center">商品总价:</td>
                                            <td  colspan="2"  class="otderpa"><input id="tprice"  value="0" size="7" name="tprice" readonly="TRUE">&nbsp;元
                                            </td>
                                        </tr>                                        <tr>
                                            <td align="center">省份/城市:</td>
                                            <td  colspan="2"  class="otderpa"><input  class="inputone" maxlength="10" size="10" name="Sheng" id="Sheng" />&nbsp;省&nbsp;/&nbsp;
                                              <input  class="inputone" maxlength="10" size="10" name="shi" id="shi" />&nbsp;市
                                              <font color="red">*</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="center">联系地址:</td>
                                            <td  colspan="2"  class="otderpa"><input class="inputone" maxlength="60" size="32" name="Addres" id="Addres" />&nbsp;<font color="red">*</font>(请填写真实联系地址）
                                            </td>
                                        </tr>
                                        <tr>
                                            <td  align="center">邮政编码:</td>
                                            <td  colspan="2"  class="otderpa"><input class="inputone" maxlength="6" size="13" name="ZipCode" id="ZipCode" />
                                                <font color="red">*</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td  align="center">付款方式:</td>
                                            <td  colspan="2"  class="otderpa">
                                                    <select class="inputone" size="1" name="HuiKuan" id="HuiKuan" onchange="jisuanpay()">
                                                        <option value="货到付款" selected="selected">货到付款</option>
                                                    </select> &nbsp;<strong>
                                        </tr>
                                        <tr>
                                            <td width="18%" align="center">验证码:</td>
                                            <td  colspan="2"  class="otderpa"><input type="text" name="check_left" id="check_left" size="6" maxlength="4" class="inputone" />&nbsp;<img src="Include/newcode_left.asp" alt="验证码看不清楚?请点击刷新验证码!" title="验证码看不清楚?请点击刷新验证码!" height="22" style="cursor:pointer;margin-bottom:-6px;" onClick="this.src='Include/newcode_left.asp?t='+(new Date().getTime());" ></td>
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
			response.Write("暂无产品信息，不能定购！")
			response.Write("<input type='hidden' name='On_RecordCount' id='On_RecordCount' value='0'>")
		else
			dim i
			i=1
    		while not rs.eof 
				if i=2 then
				response.Write("<tr>")
				end if
				response.Write("<td width=""36%"" class=""otderpa""><input class=""inputone"" maxlength=""60"" size=""32"" name=""ProductName"&i&""" id=""ProductName"&i&""" value="""&rs("ProductName")&"("&rs("Price")&"元/盒)"&""" style=""width:180px;"" readonly=""readonly"" /><font color=""red"">*</font></td>")
				response.Write("<td width=""46%"">商品数量:&nbsp;")
					response.Write("<select name='Numbers"&i&"' size='1' id='Numbers"&i&"' onchange=""jisuan()"">")
						response.Write("<option value='NULL' selected>请选择数量</option>")
						response.Write("<option value='"&rs("ProductName")&"(0)'>0盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(1)'>1盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(2)'>2盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(3)'>3盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(4)'>4盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(5)'>5盒</option>")
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
			response.Write("暂无产品信息，不能定购！")
			response.Write("<input type='hidden' name='RecordCount' id='RecordCount' value='0'>")
		else
			dim i
			i=1
    		while not rs.eof 
				response.Write(rs("ProductName")&"：")
				response.Write("<label>")
					response.Write("<select name='Numbers"&i&"' size='1' id='Numbers"&i&"'>")
						
						response.Write("<option value='NULL' selected>请选择数量</option>")
						response.Write("<option value='"&rs("ProductName")&"(0)'>零盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(1)'>一盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(2)'>二盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(3)'>三盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(4)'>四盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(5)'>五盒</option>")
					response.Write("</select>")
				response.Write("</label>")
				response.Write("【<font color='#ff0000'>"&rs("Price")&rs("PriceText")&"</font><font color='#ff0000'>免费快递送货</font>】")
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
			alert("请填写收货人信息！");
			Sh_Name.focus();
			return false;
		}
		if(ShTel.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写手机号码！");
			ShTel.focus();
			return false;
		}
		if(ShTel.value.length != "11" && ShTel.value[0] == "1" || ShTel.value.length > "13")
		{ 
			alert("请正确填写11位手机号码！");
			ShTel.focus();
			return false;
		}
		if(shi.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写市级信息！");
			shi.focus();
			return false;
		}
		if(Addres.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写收货人地址信息！");
			Addres.focus();
			return false;
		}
		if(ZipCode.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写邮编！");
			ZipCode.focus();
			return false;
		}
		if(check_left.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填验证码！");
			check_left.focus();
			return false;
		}
		if(parseInt(RecordCount.value)<=0)
		{
			alert("暂无商品信息，无法继续！")
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
				alert("请选择定购商品的数量！");
				return false;
			}
		
		}
		return true;
		}
	</script>