<!--#Include file="Head.asp"-->
  <tr><td height="10"></td></tr>
  <tr>
    <td><table width="982" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="216" valign="top"></td>
        <td width="766" valign="top"><table width="766" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><%'=Guanggao(16)%></td>
          </tr>
          <tr><td height="10"></td></tr>
          <tr>
            <td class="bk_zt pd_2"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td class="bj_9"><table border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="105" height="50" align="center" class="bj_10 txe5 fw_bd">���߶���</td>
                    <td width="637" valign="top"><table border="0" align="right" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="6" colspan="3"></td>
                        </tr>
                      <tr>
                        <td width="20" height="20" class="bj_11"></td>
                        <td class="bj_12 txe4">��ǰλ��:<a href="Index.asp" class="a3">��ҳ</a> > <a href="Order.asp" class="a3">���߶���</a></td>
                        <td width="20" class="bj_13"></td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
<%
Dim ProductName,SumProduct,Price,SalePrice
SumProduct=request.Form("SumPro")
call ProductInfo(request("ProductId"))
Function ProductInfo(Id)
 Dim rs,sql
 set rs=server.CreateObject("Adodb.recordset")
 sql="Select * from NwebCn_Products where ViewFlag and Id = "&Id&""
 rs.open sql,conn,1,1
 if not rs.eof then
  ProductName=rs("ProductName")
  Price=rs("Price")
  SalePrice=rs("Price2")
 end if
 rs.close
 set rs=nothing
End Function
%>              <tr>
                <td height="1347" valign="top" class="pd_9 lh_18">
                <%
				IF request("Action")<>"Right" then
				%>
                <style type="text/css">
<!--
.STYLE6 {	font-size: 18;
	color: #FF0000;
}
.STYLE7 {
	font-family: "����";
	font-size: 18px;
}
-->
</style>

<div class="listRight"  style="height:auto;">
  <div class="div3" style="height:auto;">
		  <table width="96%" border="0" cellspacing="0" cellpadding="0" style="margin:auto; text-align:left; margin-top:10px;margin-bottom:13px;">
           	<form name="On_Order" id="On_Order" method="post" action="AddOrder.asp" onsubmit="return Check_OnOrder();">
            <tr>
              <td width="15%" height="30">��Ʒ���ƣ�</td>
              <td height="30" colspan="3"><%=ProductName%>
              <input type='hidden' name='ProductId' id='ProductId' value='<%=ForMatDate(now,2)%>'>
              <input type='hidden' name='ProductName' id='ProductName' value='<%=ProductName%>'>
              </td>
            </tr>
            <tr>
              <td height="30">����ʱ�䣺</td>
              <td height="30" colspan="3"><%=FormatDate(Now(),1)%><input type="hidden"  name="On_dgtime" id="On_dgtime" value="<%=FormatDate(Now(),1)%>"></td> 
            </tr>
            <tr>
              <td height="30">�� �� �ˣ�</td>
              <td height="30" colspan="3"><input name="On_ShName" type="text" class="input4" size="20" id="On_ShName" />
               ������д��ʵ���������� </td>
            </tr>
            <tr>
              <td height="30">��ϵ�绰��</td>
              <td height="30" colspan="3"><input name="On_ShTel" type="text" class="input4" size="20" id="On_ShTel" /> ������д��ʵ�ջ��绰������ 
</td>
            </tr>
            <tr>
              <td height="30">&nbsp;</td>
              <td height="30" colspan="3" style="line-height:20px;"><span class="STYLE6">��ע�⣡��������Я�����ƶ��绰���룬�����ű��ڿ�ݹ�˾�ͻ�ʱ��ʱ����ȡ����ϵ��</span></td>
            </tr>
            <tr>
              <td height="30">����������</td>
              <td height="30" colspan="3"><input style="width:50px;" type='text' onblur="xx(this.value)" name='On_RecordCount' id='On_RecordCount' onkeyup="this.value=this.value.replace(/\D/g,'')" onafterpaste="this.value=this.value.replace(/\D/g,'')" value='<%=SumProduct%>'> ��/<span id="X"></span><%'=Price*SumProduct%>Ԫ (���л�֧������<%=SalePrice%>Ԫ/�� �������<%=Price%>Ԫ/��)
              <input type="hidden" name="SumMemony" id="SumMemony" value="δ֪" />
              </td>
            </tr>
           <tr>
              <td height="69">�ջ���ַ��</td>
              <td height="69" colspan="3">
              <input name="On_Sheng" type="text" class="input4" size="8" id="On_Sheng" />
              ʡ(�����ֱϽ�пɲ���)
              <input name="On_Shi" type="text" class="input4" size="8" id="On_Shi" />
              ��<br />
              <input name="On_Xian" type="text" class="input4" size="16" id="On_Xian" />
              <input type="radio" value="1" name="On_QuType" checked="checked"/>��<input type="radio" name="On_QuType" value="0" />
              �أ�����ȷѡ�������أ�<br />
              <span style="margin-left:0px;">
              <input name="On_Addres" type="text" class="input4" size="26" id="On_Addres" />
              (����д��ʵ��ϵ��ַ������
              </span>                        </td>
            </tr>
            <tr>
              <td height="30">�������룺</td>
              <td height="30" colspan="3">
              <input name="On_ZipCode" type="text" class="input4" id="On_ZipCode" size="6" maxlength="6"  /></td>
            </tr>
            <tr>
              <td height="30">�ͻ���ʽ��</td>
              <td height="30" colspan="3">��ݹ�˾�ͻ����� </td>
            </tr>
            <tr>
              <td height="30">֧����ʽ��</td>
              <td height="30" colspan="3"><span class="STYLE2"></span>
              	<div class="btn1">
                	<select name="HuiKuan" id="HuiKuan" size="1" onchange="yy(this.value)">
                    	<option selected value="��������">��������</option>
                    	 <option value="֧��������">֧��������������</option>
					     <option value="ũҵ���л��">ũҵ���л��</option>
                         <option value="�������л��">�������л��</option>
                         <option value="�������л��">�������л��</option>
                    </select><span id="Info"></span>
                </div>              </td>
            </tr>
          <script language="javascript" type="text/javascript">
		     function yy(ty){
		 		 var BPrice=<%=Price%>;
				 var SPrice=<%=SalePrice%>;
				 var Count=document.getElementById("On_RecordCount").value;
				 if(ty=="��������"){
					 xx(Count,BPrice)
					 }else{
						 xx(Count,SPrice)
						 }
				 
				 if(ty!="��������" ){
					 if(ty=="֧��������"){
							 document.getElementById("Info").innerHTML="&nbsp;��ѡ��֧����������������Ż�"+(BPrice-SPrice)*Count+"Ԫ��";
							 }else{
					 document.getElementById("Info").innerHTML="&nbsp;��ѡ��"+ty+"(���л��)���Ż�"+(BPrice-SPrice)*Count+"Ԫ��"
							 }
					 }else{
					 document.getElementById("Info").innerHTML="";
					 }
				 }
             function xx(X){
		 		 var BPrice=<%=Price%>;
				 var SPrice=<%=SalePrice%>;
				 var Count=document.getElementById("On_RecordCount").value;
				 var ty=document.getElementById("HuiKuan").value;
				 if(ty=="��������" && Count>0){
					 document.getElementById("X").innerHTML=X*BPrice;
					 document.getElementById("SumMemony").value=X*BPrice;
					 document.getElementById("Info").innerHTML=""
					 }else{
						 document.getElementById("X").innerHTML=X*SPrice;
						 document.getElementById("SumMemony").value=X*SPrice;
						 if(ty=="֧��������"){
							 document.getElementById("Info").innerHTML="&nbsp;��ѡ��֧����������������Ż�"+(BPrice-SPrice)*Count+"Ԫ��";
							 }else{
								  document.getElementById("Info").innerHTML="&nbsp;��ѡ��"+ty+"(���л��)���Ż�"+(BPrice-SPrice)*Count+"Ԫ��";
								 }
						 }
				 
				 }
            </script>

            <tr>
              <td height="30">��ע��</td>
              <td height="30" colspan="3"><textarea rows="5" style="width:90%;" name="Beizhu" id="Beizhu"></textarea></td>
            </tr>
            <tr>
              <td height="50" colspan="4" align="center" valign="bottom"><input type="submit" size="3" value="ȷ�϶���" name="Submit4" style=" width:80;font-size:15px;" class="button"/></td>
            </tr>
            </form>
            <tr>
              <td height="20" colspan="4" align="center" valign="bottom"></td>
            </tr>
            <tr>
              <td height="20" colspan="4" valign="bottom"><%
call AboutView(63)
%></td>
            </tr>
          </table>
  </div>
	  </div>
    <script language="javascript">
	<!--
	function Check_OnOrder()
	{
		var On_dgtime,On_ShName,On_ShMoble,On_ShTel,On_Sheng,On_Shi,On_Xian,On_Addres,HuiKuan,On_RecordCount;
		On_dgtime=document.getElementById("On_dgtime");
		On_ShName=document.getElementById("On_ShName");
		//On_ShMoble=document.getElementById("On_ShMoble");
		On_ShTel=document.getElementById("On_ShTel");
		On_Sheng=document.getElementById("On_Sheng");
		On_Shi=document.getElementById("On_Shi");
		On_Xian=document.getElementById("On_Xian");
		On_Addres=document.getElementById("On_Addres");
		HuiKuan=document.getElementById("HuiKuan");
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
		
			if(On_ShTel.value==""){
					alert("����д�绰���룡");
					On_ShTel.select();
					return false;
				}else{
					if(On_ShTel.value.replace(/^\s*|\s*$/g,'')!="")
					{
						var moble=On_ShTel.value;
						var patrn1=/^[+]{0,1}(\d){1,3}[ ]?([-]?((\d)|[ ]){1,12})+$/;
							if(!patrn1.exec(moble))
							{
								alert("����д��ȷ�ĵ绰���룡");
								On_ShTel.select();
								return false;
							}
					}
				}
		if(On_RecordCount.value=="" || On_RecordCount.value<1){
			alert("��������Ҫ���������!")
			On_RecordCount.focus();
			return false;
			}
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
		
		
		if(On_Sheng.value==""){
			alert("����д�����ڵ�ʡ�ݣ�")
			On_Sheng.focus();
			return false;
			}
		
		if(On_Addres.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ջ��˵�ַ��");
			On_Addres.focus();
			return false;
		}
		if(HuiKuan.value=="NULL")
		{
			alert("����ѡ�񸶿ʽ��");
			return false;
		}
		return true;
	}
	
	-->
	</script>
    



    
    
    <%end if%>
                </td>
              </tr>
            </table></td>
          </tr>
          <tr><td height="15"></td></tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="3" class="bj_7"></td>
  </tr>
  <tr>
    <td class="bj_8"><!--#Include file="Foot.asp"--></td>
  </tr>
</table>
</body>
</html>
