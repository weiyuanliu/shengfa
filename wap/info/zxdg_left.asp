<%
function getipadd()
 ipadd=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
 if ipadd= "" Then ipadd=Request.ServerVariables("REMOTE_ADDR")
 getipadd=ipadd
end function
%>
       <div style="text-align:center" id="dingou"><img src="images/index_53.jpg" /></div>
		  <table width="98%" border="0" cellspacing="0" cellpadding="0" style="margin:0 auto; text-align:left; margin-top:10px;">
            <form name="order1" id="order1" action="order.asp?Action=Left" method="post" onsubmit="return CheckOrder1(); ">
              <input type="hidden" name="ipadd" value="<%=getipadd%>" />
              <input type="hidden" name="dgtime" value="<%=Now()%>" />
			<%
			Dim THISO:THISO=str&right(year(now),1)&month(now)&day(now)&XXL(5)
			OrderId = HaveOrderId(str,THISO)
			%>
              <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>" />
              <input type="hidden" name="ZipCode" id="ZipCode" size="6" maxlength="6" value="000000"  />
            <tr>
              <td height="30" colspan="4" style="line-height:25px;">
              	<%Call ProdList()%>
              </td>
            <tr>
            <tr>
              <td width="30%" height="30">收 货 人：</td>
              <td width="70%" height="30" colspan="3"><input name="Sh_Name" type="text" class="input4" size="15" maxlength="20" id="Sh_Name" /> <font color="#ff0000">*</font></td>
            </tr>
            <tr>
              <td width="30%" height="30">手机号码：</td>
              <td width="70%" height="30" colspan="3"><input name="Sh_Tel" type="text" class="input4" size="15" id="Sh_Tel" /> <font color="#ff0000">*</font></td>
            </tr>
            <tr>
              <td width="30%" height="30">收货地址：</td>
              <td width="70%" height="30" colspan="3"><textarea name="Addres" style="height:48px; width:92%; border:1px solid #CCC;" id="Addres" align="middle" onblur="this.value=this.value.replace(/\(/g,'（');this.value=this.value.replace(/\)/g,'）')" /></textarea> <font color="#ff0000">*</font></td>
            </tr>
            <tr>
              <td width="30%" height="30"></td>
              <td width="70%" height="30" colspan="3"><font color="#ff0000">请填写详细收货地址（如：XX省XX市XX区XX街道XX号）</font></td>
            </tr>
            <tr>
              <td height="30" colspan="4" align="center" valign="bottom"><input id="cod" name="tijiao" type="image" src="images/order_18.png" style="border:0;" /></td>
            </tr>
            </form>
          </table>
  </div>
      
<%
    sub ProdList()
    	dim rs,sql,Pml
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
				response.Write("<label>")
				response.Write("<strong>"&rs("ProductName")&"</strong>")

				response.Write("&nbsp;<font color='#ff0000'>"&rs("Price")&rs("PriceText")&"</font>&nbsp;")

					response.Write("<select name='Numbers"&i&"' size='1' id='Numbers"&i&"'>")
						response.Write("<option value='NULL' selected>选择订购数量</option>")
						response.Write("<option value='"&rs("ProductName")&"(0)'>0盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(1)'>1盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(2)'>2盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(3)'>3盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(4)'>4盒</option>")
						response.Write("<option value='"&rs("ProductName")&"(5)'>5盒</option>")
					response.Write("</select>")
				response.Write("</label>")
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
	<!--
		function CheckOrder1()
		{
			var dgtime,Sh_Name,Sh_Mobel,Sh_Tel,Addres,RecordCount,check_left
			Sh_Name=document.getElementById("Sh_Name");
			ShTel=document.getElementById("Sh_Tel");
			Addres=document.getElementById("Addres");
			check_left=document.getElementById("check_left");
			RecordCount=document.getElementById("RecordCount");
			
		if(Sh_Name.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写收货人信息！");
			Sh_Name.focus();
			return false;
		}		var telnum=/[0-9-]+$/
		if(isNaN(ShTel.value))
		{
			alert("手机号码必须为数字！");
			ShTel.focus();
			return false;
		}
		if(ShTel.value.length != "11" && ShTel.value[0] == "1" || ShTel.value.length > "13")
		{ 
			alert("请正确填写11位手机号码！");
			ShTel.focus();
			return false;
		}
		if(Addres.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("请填写收货人地址信息！");
			Addres.focus();
			return false;
		}
		if(Addres.value.length<6)
		{
			alert("请填写详细收货地址！");
			Addres.focus();
			return false;
		}		if(parseInt(RecordCount.value)<=0)
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
	-->
	</script>