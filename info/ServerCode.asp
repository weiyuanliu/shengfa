<%Sub SearchList()%>
	<%
		response.Write("OK")
		response.End()
	%>
    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6C6C6C">
        <tr>
            <td width="19%" height="28" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">订 单 编 号</span></td>
            <td width="20%" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">订货人姓名</span></td>
            <td width="22%" height="28" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">下 单 时 间</span></td>
            <td width="26%" height="28" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">定 单 状 态</span></td>
            <td width="13%" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">定 单 金 额</span></td>
        </tr>
   			<%Call SearchList(20)%>
    </table>
<%End Sub%>

<%
Sub SearchList(Page_Size)
	Dim KeyWord
	KeyWord=Trim(Request("KeyWord"))
	Dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	if KeyWord<>"" then
		sql="select id,ProductNo,Linkman,AddTime,State,HuoDao_FuKuan,Remark,FuKuan,FaHuoTime from NwebCn_Order where Linkman ='"&KeyWord&"' order by AddTime desc"
	else
		response.Write("<script language=javascript>"&vbcrlf)
			response.Write("alert('请输入查寻用户名！');"&vbcrlf)
			response.Write("window.history.go(-1);")
		response.Write("</script>")
		response.End()
		exit sub
	end if
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr>")
			response.Write("<td colspan='5' align='left' style='padding:5px;'>")
				response.Write("对不起，暂没有找你要的信息！")
			response.Write("</td>")
		response.Write("</tr>")
	else
		rs.pagesize=Page_Size
		dim sum_page,total,i
		total=rs.recordcount
		sum_page=total \ page_size
		if total mod page_size <>0 then sum_page=sum_page+1
		dim page
		page=trim(request.querystring("page"))
		if page="" or isnull(page) or (not IsNumeric(page)) then
			page=1
		elseif Cint(Page)<=1 then
			page=1
		elseif Cint(page) => sum_page then
			page=sum_page
		else
			page=Cint(page)
		end if
		rs.absolutepage=page
		dim Flage
		 
		for i=1 to Page_Size
			if not rs.eof then
			flage=1
				response.Write("<tr bgcolor='#444444' height='25'>")
					response.Write("<td>")
						response.Write("<a href='OrderView.asp?ID="&rs("id")&"' target='_blank'>")
						response.Write(rs("ProductNo"))
						response.Write("</a>")
					response.Write("</td>")
					response.Write("<td>")
						response.Write(rs("Linkman"))
					response.Write("</td>")
					response.Write("<td>")
						response.Write(rs("AddTime"))
					response.Write("</td>")
					response.Write("<td>")
						if rs("State")<>"" then 
							'response.Write(rs("State"))
						else
							response.Write("待处理……")						
						end if
						if INstr(Rs("State"),"未付款")>0 Then
						Response.Write("<br><font color='#ff0000'>您选择的是先付款（银行汇款或者支付宝付款），但是我们至今未收到您的货款，所以没有发货。请及时付款并在付款后通知我们，我们就会及时为您发货。如果您还有问题请打400-661-9898咨询</font>")
						flage=0
						end if
						if INstr(Rs("State"),"不能到付")>0 Then
						Response.Write("<br><font color='#ff0000'>对不起，您的收货地没有能够代收货款的快递公司，所以不能发货，如果您还需要我们的产品，请重新提交银行汇款（或者支付宝付款）的订单，我们收到您付的货款后就会发货，如果您还有问题请打400-661-9898咨询。</font>")
						flage=0
						end if
						if instr(Rs("State"),"钱到已发")>0 then
					'	response.Write("<font color='#ff0000'>"&rs("State")&"<br>")   

									 response.Write("&nbsp;&nbsp;您的货已经发出（如果6天后未收到货请联系我们）， 发货时间："&FormatDate(rs("FaHuoTime"),4))
									 Response.Write("。如果您还有问题请打400-661-9898咨询。</font>")
									 flage=0
						end if
						
						
						if instr(Rs("State"),"已经发货")>0 then
					'	response.Write("<font color='#ff0000'>"&rs("State")&"<br>")
									 response.Write("&nbsp;&nbsp;您的货已经发出（如果6天后未收到货请联系我们）， 发货时间："&FormatDate(rs("FaHuoTime"),4)) 
									 Response.Write("。如果您还有问题请打400-661-9898咨询。</font>")
									 flage=0
						end if
						if instr(Rs("State"),"刚订未发")>0 then
						Response.Write("<font color='#ff0000'>我们会尽快处理您的订单（我们接到提交的订单，处理时限为不超过12个小时），请稍后再查询。如果您还有问题请打400-661-9898咨询。</font>")
						flage=0
						end if
						if instr(Rs("State"),"未处理")>0 then
						Response.Write("<font color='#ff0000'>我们会尽快处理您的订单（我们接到提交的订单，处理时限为不超过12个小时），请稍后再查询。如果您还有问题请打400-661-9898咨询。</font>")
						flage=0
						end if
						if flage=1 then
						Response.Write("<font color='#ff0000'>我们会尽快处理您的订单（我们接到提交的订单，处理时限为不超过12个小时），请稍后再查询。如果您还有问题请打400-661-9898咨询。</font>")
						end if
					'if rs("HuoDao_FuKuan") then
						'if rs("FuKuan") then
							'if rs("State")="货到后付款" then
								'response.Write("等待发货……")
							'else
								'if Instr(rs("State"),"货已发")>0 then
									'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
									'response.Write("&nbsp;&nbsp;发货时间："&FormatDate(rs("FaHuoTime"),4))
								'else
									'response.Write(rs("State"))
								'end if
							'end if
						'else
							'if rs("State")="" or isnull(rs("State")) then
								'response.Write("待处理……")
							'else
							'response.Write("对不起，当地不能货到付款，货没有发！")
							'end if
						'end if
					'else
						'if Instr(rs("State"),"货已发")>0 then
							'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
							'response.Write("&nbsp;&nbsp;发货时间："&FormatDate(rs("FaHuoTime"),4))
						'else
						'Response.Write("因你的收货地不支持货到付款，请重下订单先汇款才可以发货，如果有问题请电话咨询我们。 ")
							'response.Write(rs("State"))
						'end if
					'end if
					
					response.Write("</td>")
					
					response.Write("<td>")
						response.Write(SumMemony(rs("Remark")))
					response.Write("</td>")
				response.Write("</tr>")
				rs.movenext
			end if
		next
		if sum_page>1 then call Contrl_Page(page,sum_page,total,page_size) 
	end if
End sub
%>