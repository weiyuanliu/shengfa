<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE3 {color: #FF0000; font-size: 14; }
.STYLE5 {
	color: #FF0000;
	font-size: 18px;
	font-family: "����";
}
.STYLE6 {
	font-size: 18;
	color: #FF0000;
}
-->
</style>





<div class="listLeft" style="height:auto;">
  <div class="div3" style="height:auto;">
		  <table width="98%" border="0" cellspacing="0" cellpadding="0" style="margin:auto; text-align:left; margin-top:10px; margin-bottom:20px;">
            <form name="order1" id="order1" action="order.asp?Action=Left" method="post" onsubmit="return CheckOrder1(); ">
            <tr>
              <td width="120" height="30">��Ʒ���ƣ�</td>
              <td height="30" colspan="3">����� <span class="STYLE5">������������ڴ���������</span></td>
            </tr>
            <tr>
              <td height="30">����ʱ�䣺</td>
              <td height="30" colspan="3"><%=FormatDate(Now(),4)%><input type="hidden" name="dgtime" value="<%=Now()%>">
              <%
			Dim THISO:THISO=str&right(year(now),1)&month(now)&day(now)&XXL(5)
			OrderId = HaveOrderId(str,THISO)
			%>
              <input type="hidden"  name="OrderId" id="OrderId" value="<%=OrderId%>"></td> 
            </tr>
            <tr>
              <td height="30">�� �� �ˣ�</td>
              <td height="30" colspan="3"><input name="Sh_Name" type="text" class="input4" size="20" id="Sh_Name" />
              ������д��ʵ���������� </td>
            </tr>
            <!--<tr>
              <td height="30">�� &nbsp;&nbsp;&nbsp;&nbsp; ����</td>
              <td width="33%" height="30"><input name="Sh_Mobel" type="text" class="input4" size="15" id="Sh_Mobel" /></td>
              <td width="6%" height="30">&nbsp;</td>
              <td width="46%" height="30">&nbsp;</td>
            </tr>-->
            <tr>
              <td height="30">��ϵ�绰��</td>
              <td height="30" colspan="3"><input name="Sh_Tel" type="text" class="input4" size="15" id="Sh_Tel" /></td>
            </tr>
            <tr>
              <td height="30">&nbsp;</td>
              <td height="30" colspan="3" style="line-height:20px;"><span class="STYLE6">��ע�⣡��������Я�����ƶ��绰���룬�磺�ֻ�����С����ͨ�������ű��ڿ�ݹ�˾�ͻ�ʱ��ʱ����ȡ����ϵ��</span></td>
            </tr>
            <tr>
              <td height="30">����������</td>
              <td height="30" colspan="3">&nbsp;</td>
            </tr>
            <tr>
              <td height="30" colspan="4" style="line-height:25px; padding-left:20px;">
              	<%Call ProdList()%>              </td>
            <tr>
              <td height="30" colspan="4" align="center"><span class="STYLE1">ע����ѿ���ͻ�����</span></td>
            </tr>
            <tr>
              <td height="30">�ջ���ַ��</td>
              <td height="30" colspan="3"><input name="Sheng" type="text" class="input4" id="Sheng" size="8" />
              ʡ(�����ֱϽ�пɲ���)
              <input name="shi" type="text" class="input4" id="shi" size="8" />
              ��<br />
              <input name="xian" type="text" class="input4" id="xian" size="16" />
              <input type="radio" value="1" name="QuType" id="QuType" checked="checked"/>��<input type="radio" name="QuType" id="QuType" value="0" />
              ��<span style="margin-left:0px;">������ȷѡ�������أ�<br />
              <input name="Addres" type="text" class="input4" id="Addres" size="26" onkeyup="this.value=this.value.replace('(','��');this.value=this.value.replace(')','��')" />
              ������д��ʵ��ϵ��ַ������
              </span></td>
            </tr>
            <tr>
              <td height="30">�������룺</td>
              <td height="30" colspan="3"><input type="text" name="ZipCode" id="ZipCode" size="6" maxlength="6" class="input4"  />              </td>
            </tr>
            <tr>
              <td height="30">�ͻ���ʽ��</td>
              <td height="30" colspan="3">��ѿ���ͻ�����</td>
            </tr>
            <tr>
              <td height="30">֧����ʽ��</td>
              <td height="30" colspan="3">�������� (<span class="STYLE3"><font color="#FF0000">���л�����Ż�</font></span>)</td>
            </tr>
            <tr>
              <td height="80" colspan="4" align="center" valign="bottom"><input id="cod" name="tijiao" type="image" src="images/btn05.jpg"  /></td>
            </tr>
            </form>
          </table>
  </div>
	  </div>
      
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
	<!--
		function CheckOrder1()
		{
			var dgtime,Sh_Name,Sh_Mobel,Sh_Tel,Sheng,shi,xian,Addres,RecordCount,ZipCode
			Sh_Name=document.getElementById("Sh_Name");
			//Sh_Mobel=document.getElementById("Sh_Mobel");
			ShTel=document.getElementById("Sh_Tel");
			//Sheng=document.getElementById("Sheng");
			shi=document.getElementById("shi");
			xian=document.getElementById("xian");
			Addres=document.getElementById("Addres");
			ZipCode=document.getElementById("ZipCode");
			RecordCount=document.getElementById("RecordCount");
			
		if(Sh_Name.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ջ�����Ϣ��");
			Sh_Name.focus();
			return false;
		}
		
		
		/*if(Sheng.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����дʡ����Ϣ��");
			Sheng.focus();
			return false;
		}*/
		
		if(shi.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�м���Ϣ��");
			shi.focus();
			return false;
		}
		if(xian.value.replace(/^\s*|\s*$/g,'')=="")
		{
			alert("����д�ؼ���Ϣ��");
			xian.focus();
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
	-->
	</script>