<!--#Include file="Head.asp"-->
<!--#include file="../info/ProdClass.asp"-->
<%
Dim Prod
Set Prod=New ProdClass
Prod.Set_TextID(10)
Prod.Set_TableName("NwebCn_About")
%>
	</div>
<div id="main">
  <div class="topad1"><img src="images/io_tops.jpg" /></div>
  <div class="html">
  <div class="leftnews">
     <div style="color:#B80F07;font-size:16px;font-weight:bold;  border-bottom:1px solid #B80F07;  height:30px; max-width:500px">��������</div>
          <div class="leftnews1">
          <div class="listnews">
          <ul>
                           <%
			  call ArticleList("0,58,")
			  Function ArticleList(SPath)
			  Dim Rs,Sql
			  Set Rs=Server.CreateObject("Adodb.Recordset")
			  Sql="Select * from NwebCn_News where ViewFlag=1 and Charindex(SortPath,'"&Spath&"')>0 order by px,id desc"
			  Rs.open Sql,conn,1,3
			  If Rs.Eof Then
			  Response.Write("<li>��ʱû�������Ϣ</li>")
			  Else
					dim page,sum_count,pagescount
					rs.pagesize=25
					sum_count=rs.recordcount
					pagescount=sum_count \ rs.pagesize
					if sum_count mod rs.pagesize <>0 then pagescount=pagescount+1
					page=trim(request.QueryString("page"))
					if page="" or isnull(page) or (not IsNumeric(page)) then
						page=1
					elseif Cint(page)<=1 then
						page=1
					elseif Cint(page)>pagescount then
						page=pagescount
					else
						page=Cint(page)
					end if
					rs.absolutepage=page
					Dim ii,jj
			 %>
		<%For ii=1 to rs.pagesize/1%>
		<%
		For jj=1 to 1
		If Not Rs.eof Then
		%>
          <li><a href="NewsView.asp?Id=<%=rs("Id")%>&SortId=<%=rs("SortId")%>" title="<%=rs("NewsName")%>" target="_blank"><%=rs("NewsName")%></a></li>
		<%
		Rs.MoveNext
		End if
		Next
		%>
		<%Next%>
          </ul>
	<%if rs.recordcount>rs.pagesize then%>
          <div class="page"><DIV class="pagelistbox"><span><%=dispartpage(page,pagescount,"News.asp")%></span></DIV></DIV>
	<%end if%>
	<%
	End If
	Rs.Close
	Set Rs=Nothing
	End Function
	%>
          </div>
          </div>
     </div>
     <div>
  </div>
  </div>
  </div>
<!--#Include file="Foot.asp"-->