<!--#Include file="Head.asp"-->
     <div class="topad"></div>
     <div style="background:url(style/blue/images/datu.gif) no-repeat center; margin:0 auto; width:1420px; height:410px;"><div style="width:1420px;height:568px;position:absolute;"></div></div>
     <div style="background:url(style/blue/images/home_top_11.png) no-repeat center; margin:0 auto; width:1101px; height:271px;"></div>
	</div>
<div id="main"class="loads">
  <div><p style="text-align:center;" ><img src="style/blue/images/homea_08.png" width="1420" height="556" /></p>
  <p><img src="style/blue/images/homea_14.png" width="1420" height="757" /></p>
  <p style="text-align:center;" ><img src="style/blue/images/homea_16.png" width="1420" height="543" /></p>
  <p style="text-align:center;" ><img src="style/blue/images/fengexian.png" width="1420" height="343" /></p>
  <p style="text-align:center;" ><img src="style/blue/images/homeb_04.png" width="1420" height="356" /></p>
   <p style="text-align:center;" ><img src="style/blue/images/miaoshu5.png" width="1420" height="547" /></p>
  <p style="text-align:center;" ><img src="style/blue/images/homer_01.jpg" width="1420" height="687" /></p>
  <p style="text-align:center;" ><img src="style/blue/images/miaoshu7.gif" width="1420" height="587" /></p>
  <p style="text-align:center;" ><img src="style/blue/images/homer_02.png" width="1420" height="535" /></p>
  <p><img src="style/blue/images/homer_03.png" width="1420" height="431" /></p>
  <p><img src="style/blue/images/fengexian.png" width="1420" height="343" /></p>
  <p><img src="style/blue/images/homer_05.png" width="1420" height="512" /></p>
  <p><img src="style/blue/images/homer_06.png" width="1420" height="254" /></p>
  <p><img src="style/blue/images/homer_07.png" width="1420" height="95" /></p>
  <p><img src="style/blue/images/homer_08.jpg" width="1420" height="744" /></p>
  <p><img src="style/blue/images/homer_09.png" width="1420" height="496" /></p>
  <p><img src="style/blue/images/homer_10.png" width="1420" height="412" /></p>
  <p><img src="style/blue/images/homer_11.jpg" width="1420" height="432" /></p>
  <p><img src="style/blue/images/homer_12.png" width="1420" height="675" /></p>
  <p><img src="style/blue/images/homer_13.png" width="1420" height="540" /></p>
  <p><img src="style/blue/images/homer_14.png" width="1420" height="478" /></p>
  <p><img src="style/blue/images/homer_15.png" width="1420" height="629" /></p>
  <p><img src="style/blue/images/homer_16.png" width="1420" height="192" /></p>
  
  
  
  
  <p><img src="style/blue/images/homer_42.png" width="1420" height="116"  /></p>
    <div style=" background:url(style/blue/images/homed_13.jp) no-repeat; height:783px"><!----留言背景----->
      <div style="float:left;width:750px;height:730px;overflow:hidden;"><div style="padding:15px 0 0 220px">
<div class="homefeedbackl">
        <!----feedback----->
	<%
	if sum_count > rs.pagesize then
	%>
	<%
	end if
	%>
         		<div style="height:730px; overflow-y:scroll;">
                <div class="feedbacktable">
             <%
			Function DeleteMessage()
			 dim sql
			 sql="delete from NwebCn_Message where Addtime < #2010-06-01# or instr(Content,'http:')>0 or instr(MesName,'taobao')>0"
			 conn.execute(sql)
			End Function
			call MessageList()
			Function MessageList()
			 Dim rs,sql
			 set rs=server.CreateObject("Adodb.recordset")
			 sql="Select * from NwebCn_Message where SecretFlag=1 and ViewFlag=1 Order by AddTime desc"
			 rs.open sql,conn,1,1
			 if not rs.eof then
			 dim page,sum_count,pagescount
					rs.pagesize=30
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
					Dim ii,jj,Ip,Ix,i,IM
				For ii=1 to rs.pagesize/1%>
            <%
			 For jj=1 to 1
			 If Not Rs.eof Then
			 Ip=rs("Mobile")
			 Ix=split(Ip,".")
			 Ip=Ix(0)&"."&Ix(1)&"."&Ix(2)&".**"
			%>
                  <ul>  
                     <li><span>来&nbsp;&nbsp;自：</span><font style="width:140px;"><%=rs("LinkMan")%></font><label>IP：<%=Ip%>&nbsp;日期：<%=ForMatDate(rs("AddTime"),14)%></label></li>
                     <b><font color="#b80f07"><li><span>主&nbsp;&nbsp;题：</span><font style="width:400px;"><%=RS("MesName")%></font></li></font></b>
                     <li><span>留&nbsp;&nbsp;言：</span><font style="width:400px;"><%=Replace(Replace(rs("Content"),"&lt;br&gt;",""),"&nbsp;"," ")%></font></li>
		<%if rs("ReplyContent")<>"" then%>
                     <li class="reds"><span>回&nbsp;&nbsp;复：</span class="reds"><font style="width:400px;"><%=Replace(rs("ReplyContent"),"&nbsp;&lt;br&gt;","")%></font></li>
		<%end if%>
                  </ul>
		<%
		Rs.MoveNext
		End if
		Next
		%>
		<%Next%>
	<%
	end if
	rs.close
	set rs=nothing
	End Function
	%>
               </div>
            </div>
        <!----feedback end----->
        </div>
      </div></div>
      <div style="float:left;margin-left:65px;width:360px;height:700px;overflow:hidden;">
       <div class="homeorderlist">
           <ul>
<!--#include file="info/index_MsgClass.asp"-->
         <%
		 	Dim Action
			Action=Trim(Request("Action"))
			Select Case Action
				Case "Search":
					SearchKeyList
				Case "Search_Phone":
					ViewText	
				Case Else
	%>            
			<%
                            DIm Object
                            Set Object=New ViewClass
                            Object.Set_Page_Size(24)
                            Object.ViewList
			%>
	<%
	End Select
	%>
           </ul>
         </div>
     </div>
     </div>
   </div>
<!--#Include file="Foot.asp"-->