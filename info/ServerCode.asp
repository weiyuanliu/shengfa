<%Sub SearchList()%>
	<%
		response.Write("OK")
		response.End()
	%>
    <table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6C6C6C">
        <tr>
            <td width="19%" height="28" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">�� �� �� ��</span></td>
            <td width="20%" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">����������</span></td>
            <td width="22%" height="28" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">�� �� ʱ ��</span></td>
            <td width="26%" height="28" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">�� �� ״ ̬</span></td>
            <td width="13%" align="center" bgcolor="#1B1B1B"><span style="font-weight: bold">�� �� �� ��</span></td>
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
			response.Write("alert('�������Ѱ�û�����');"&vbcrlf)
			response.Write("window.history.go(-1);")
		response.Write("</script>")
		response.End()
		exit sub
	end if
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.Write("<tr>")
			response.Write("<td colspan='5' align='left' style='padding:5px;'>")
				response.Write("�Բ�����û������Ҫ����Ϣ��")
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
							response.Write("��������")						
						end if
						if INstr(Rs("State"),"δ����")>0 Then
						Response.Write("<br><font color='#ff0000'>��ѡ������ȸ�����л�����֧���������������������δ�յ����Ļ������û�з������뼰ʱ����ڸ����֪ͨ���ǣ����ǾͻἰʱΪ������������������������400-661-9898��ѯ</font>")
						flage=0
						end if
						if INstr(Rs("State"),"���ܵ���")>0 Then
						Response.Write("<br><font color='#ff0000'>�Բ��������ջ���û���ܹ����ջ���Ŀ�ݹ�˾�����Բ��ܷ��������������Ҫ���ǵĲ�Ʒ���������ύ���л�����֧��������Ķ����������յ������Ļ����ͻᷢ��������������������400-661-9898��ѯ��</font>")
						flage=0
						end if
						if instr(Rs("State"),"Ǯ���ѷ�")>0 then
					'	response.Write("<font color='#ff0000'>"&rs("State")&"<br>")   

									 response.Write("&nbsp;&nbsp;���Ļ��Ѿ����������6���δ�յ�������ϵ���ǣ��� ����ʱ�䣺"&FormatDate(rs("FaHuoTime"),4))
									 Response.Write("������������������400-661-9898��ѯ��</font>")
									 flage=0
						end if
						
						
						if instr(Rs("State"),"�Ѿ�����")>0 then
					'	response.Write("<font color='#ff0000'>"&rs("State")&"<br>")
									 response.Write("&nbsp;&nbsp;���Ļ��Ѿ����������6���δ�յ�������ϵ���ǣ��� ����ʱ�䣺"&FormatDate(rs("FaHuoTime"),4)) 
									 Response.Write("������������������400-661-9898��ѯ��</font>")
									 flage=0
						end if
						if instr(Rs("State"),"�ն�δ��")>0 then
						Response.Write("<font color='#ff0000'>���ǻᾡ�촦�����Ķ��������ǽӵ��ύ�Ķ���������ʱ��Ϊ������12��Сʱ�������Ժ��ٲ�ѯ������������������400-661-9898��ѯ��</font>")
						flage=0
						end if
						if instr(Rs("State"),"δ����")>0 then
						Response.Write("<font color='#ff0000'>���ǻᾡ�촦�����Ķ��������ǽӵ��ύ�Ķ���������ʱ��Ϊ������12��Сʱ�������Ժ��ٲ�ѯ������������������400-661-9898��ѯ��</font>")
						flage=0
						end if
						if flage=1 then
						Response.Write("<font color='#ff0000'>���ǻᾡ�촦�����Ķ��������ǽӵ��ύ�Ķ���������ʱ��Ϊ������12��Сʱ�������Ժ��ٲ�ѯ������������������400-661-9898��ѯ��</font>")
						end if
					'if rs("HuoDao_FuKuan") then
						'if rs("FuKuan") then
							'if rs("State")="�����󸶿�" then
								'response.Write("�ȴ���������")
							'else
								'if Instr(rs("State"),"���ѷ�")>0 then
									'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
									'response.Write("&nbsp;&nbsp;����ʱ�䣺"&FormatDate(rs("FaHuoTime"),4))
								'else
									'response.Write(rs("State"))
								'end if
							'end if
						'else
							'if rs("State")="" or isnull(rs("State")) then
								'response.Write("��������")
							'else
							'response.Write("�Բ��𣬵��ز��ܻ��������û�з���")
							'end if
						'end if
					'else
						'if Instr(rs("State"),"���ѷ�")>0 then
							'response.Write("<font color='#ff0000'>"&rs("State")&"</font>")
							'response.Write("&nbsp;&nbsp;����ʱ�䣺"&FormatDate(rs("FaHuoTime"),4))
						'else
						'Response.Write("������ջ��ز�֧�ֻ�����������¶����Ȼ��ſ��Է����������������绰��ѯ���ǡ� ")
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