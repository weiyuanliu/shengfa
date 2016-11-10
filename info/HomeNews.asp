<%
Dim News
Set News=New NewsClass
News.Set_TuiJian(8)
News.Set_SortPath("0,58,")
%>
<div class="right">
	  <a href="news.asp" target="_blank" ><img src="images/title01.jpg" class="imgClear"  border="0"/></a>
		<ul>
		  <%=News.ScrollTuiJianList%>
		</ul>
		<!--<img src="images/img02.jpg" class="imgClear" />-->
	  </div>