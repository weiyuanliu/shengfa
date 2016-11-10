<%
Dim Products
set Products =New ProdClass
Products.Set_TuiJian(2)
%>
<div class="left">
	 <ul>
		<%=Products.ProdTuiJian%>	 
	</ul>
 </div>