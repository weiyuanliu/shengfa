<!--#Include file="Head.asp"-->
<!--#include file="../info/NewsClass.asp"-->
<%
Dim News',ID
ID = SafeRequest("id","get")
If Not IsNum(ID) OR IsNul(ID) Then
	Put "<script>alert('数据出错，请返回！');window.history.go(-1);</script>"
End If
Set News=New NewsClass
News.Set_ID(ID)
%>

	</div>
   <div id="main">
   <div class="topad1"><img src="images/news_top.jpg" /></div>
   <div class="html">
       <div class="htmlbox">
       <div class="listrightt">你的位置： 新闻中心 &gt; 新闻查看</div><div class="listrightc"></div>
          <div class="html_con">
              <div class="html_con1"><h3><%=News.Title%></h3></div>
                <div class="blank_b"></div>
                <div class="html_content"><%=News.Text%></div>
           </div>
       </div>
    </div>
  </div>
  <div class="blank_b"></div>
<!--#Include file="Foot.asp"-->