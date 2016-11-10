<%
Class ProdClass
	Dim Rs,Sql,TuiJian,TextID,TableName,Picture_Width,Picture_Height
	
	Private Sub SetRs()
		set Rs=server.CreateObject("adodb.Recordset")
	End Sub
	Public Sub SetPicture_Width(Values)
		Picture_Width=Values
	End Sub
	Public Sub SetPicture_Height(Values)
		Picture_Height=Values
	End Sub
	Public Sub Set_TextID(Values)
		TextID=Values
	End Sub
	Public Sub Set_TableName(Values)
		TableName=Values
	End Sub
	
	Private Sub CloseRs()
		Rs.Close()
		set Rs=Nothing
	End sub
	
	Public Sub Set_TuiJian(Values)
		TuiJian=Values
	End Sub
	
	Public Sub ProdTuiJian()
		SetRs
		Sql="select top "&TuiJian&" ID,ProductName,Price,PriceText,SmallPic,BigPic from NwebCn_Products where ViewFlag=1 order by px desc,AddTime desc"
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("<li>")
				response.Write("暂无信息！")
			response.Write("</li>")
		else
			while not rs.eof
				response.Write("<li>")
					if rs("SmallPic")<>"" then
						response.Write("<img src='"&replace(rs("SmallPic"),"../","")&"' class='imgClear' onload='javascript:DrawImage(this,200,200);' />")
					elseif rs("BigPic")<>"" then
						response.Write("<img src='"&replace(rs("BigPic"),"../","")&"' class='imgClear' onload='javascript:DrawImage(this,200,200);'/>")
					else
						response.Write("<img src='Images/NoPicture.jpg' class='imgClear' onload='javascript:DrawImage(this,200,200);'/>")
					end if
					response.Write("<span onmouseover=""this.style.cursor='pointer';"" title='"&rs("Price")&rs("PriceText")&"'>价格："&rs("Price")&rs("PriceText")&"</span>")
					response.Write("<a href='zxdg.asp' title='"&rs("Price")&rs("PriceText")&"'>")
					response.Write("<img src='images/btn02.jpg' class='imgClear'/>")
					response.Write("</a>")
				response.Write("</li>")
				rs.movenext
			wend
		end if
		CloseRs()
	End sub
	
	Public Sub Text() '产品简介信息
		SetRs
		Sql="Select Content from "&TableName&" where id="&TextID&" and ViewFlag=1"
		Rs.open Sql,conn,1,1
		if rs.eof and rs.bof then
			response.Write("产品简介信息添加，请稍后……")
		else
			response.Write(rs("Content"))
		end if
		CloseRs
	End sub
	
	Public Sub ProdList()
		SetRs
		Sql="Select ID,ProductName,Price,PriceText,SmallPic,BigPic from NwebCn_Products order by px desc,AddTime desc"
		Rs.open Sql,conn,1,1
		if Not rs.eof and Not rs.bof then
			dim i,str
			i=1
			response.Write("<table border=0 cellpadding=0 cellspacing=0 wdith='618'>")
				response.Write("<tr>")
					while not rs.eof
						response.Write("<td width='309' align='center'>")
							response.Write("<table border=0 cellpadding=0 cellspacing=0 align='center'>")
								response.Write("<tr>")
									response.Write("<td>")
										if rs("SmallPic")<>"" then
											str=rs("SmallPic")
										elseif rs("BigPic")<>"" then
											str=rs("BigPic")
										else
											str="Images/NoPicture.jpg"
										end if
										response.Write("<img src='"&replace(str,"../","")&"' border=0 onload='javascript:DrawImage(this,"&Picture_Width&","&Picture_Height&");'")
									response.Write("</td>")
								response.Write("</tr>")
								response.Write("<tr>")
									response.Write("<td style='text-align:center; padding-top:5px; padding-bottom:2px;'>")
										response.Write(rs("ProductName"))
									response.Write("</td>")
								response.Write("</tr>")
								response.Write("<tr>")
									response.Write("<td style='text-align:center; padding-top:0px; padding-bottom:2px;'>")
										response.Write(rs("Price"))
										response.Write(rs("PriceText"))
									response.Write("</td>")
								response.Write("</tr>")
								response.Write("<tr>")
									response.Write("<td style='text-align:center; padding-top:0px; padding-bottom:20px;'>")
										response.Write("<a href='Order.asp?Id="&rs(0)&"'>")
										response.Write("<img src='images/btn02.jpg' class='imgClear'/>")
										response.Write("</a>")
									response.Write("</td>")
								response.Write("</tr>")
							response.Write("</table>")
						response.Write("</td>")
						if i mod 2 =0 then response.Write("</tr>")
						rs.movenext
						i=i+1
					wend
			response.Write("</table>")
		End if
		CloseRs
	End Sub
End Class
%>