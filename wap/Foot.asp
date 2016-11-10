
<%
Dim dgurl
if nm = "index" or nm = "order" or nm = "alipay" then
	dgurl="#dingou"
else
	dgurl="Order.asp"
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
  <tr>
    <td style="background:url(images/index_52.jpg);background-size:100%; color:#000000" valign="middle" align="center" height="150">
        郑重声明：未经授权禁止转载、摘编、复制或建立镜像<br />版权 &copy; 汉方演绎中国官方网站  粤ICP备14001626号

    </td>
  </tr>
</table>
    </div>
</div>
<% if  nm <> "order" then%>
<div id="ccidapDiv1">
<div align="center">
       <table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tbody><tr>
    <td style="text-align:center"><a href="<%=dgurl%>"><img src="images/btn_01.jpg" width="156" /></a></td>
    <td style="text-align:right" width="39"><a href="#top"><img src="images/top.jpg" width="39" /></a></td>
  </tr>
</tbody></table>
  </div>
</div>
<%end if%>
 

</body>
</html>

<%
Function addurl(lurl)   
domext = "com,net,org,cn,la,cc,info,hk,biz,memo,bin,am,etv,asi,ak,rde,org.cn,co.kr,com.cn,net.cn,gov.cn,com.hk" 
arrdom = Split(domext, ",")  
addurl = "": lurl = LCase(lurl)  
If lurl = "" Or Len(lurl) = 0 Then Exit Function 
lurl = Replace(Replace(lurl, "http://", ""), "https://", "")  
ds1 = InStr(lurl, ":") - 1 '过滤掉端口  
If ds1 < 0 Then ds1 = InStr(lurl, "/") - 1 '过滤掉/后面的字符  
If ds1 > 0 Then lurl = Left(lurl, ds1)  
ds2 = Split(lurl, ".")(UBound(Split(lurl, ".")))  
If InStr(domext, ds2) = 0 Then 
    addurl = lurl  
Else 
    For dd = 0 To UBound(arrdom)  
        If InStr(lurl, "." & arrdom(dd)) > 0 Then 
            addurl = Replace(lurl, "." & arrdom(dd) & "", "")  
            If InStr(addurl, ".") = 0 Then 
            addurl = lurl  
            Else 
            addurl = Split(addurl, ".")(UBound(Split(addurl, "."))) & "." & arrdom(dd)  
            End If 
        End If 
    Next 
End If 
End Function

Dim advlink,userip,advlinks,lailuy,lailu,hostlailu,lurl,flurl
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If userip = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

lailuy = Request.ServerVariables("HTTP_REFERER")
If InStr(lailuy,"?") >0 Then
	lailu = Split(lailuy,"?")(0)
else
	lailu = lailuy
end if
lurl = lailu
flurl = addurl(lurl)
if flurl = "baidu.com" then
	advlink = "www.baidu.com"
elseif flurl = "sogou.com" then
	advlink = "www.sogou.com"
elseif flurl = "haosou.com" then
	advlink = "www.haosou.com"
elseif flurl = "soso.com" then
	advlink = "www.soso.com"
elseif flurl = "sm.cn" then
	advlink = "m.sm.cn"
else
	advlink = lailu
end if

if lailuy = "" then
	hostlailu = request.ServerVariables("HTTP_HOST")
If InStr(hostlailu,"?") >0 Then
	advlink = Split(hostlailu,"?")(0)
else
	advlink = hostlailu
end if
end if


lailu = Request.ServerVariables("HTTP_REFERER")	

session("advlink")=lailu

advlink = session("advlink")
'response.write  session("advlink")
if advlink = "" then
	advlink = request.ServerVariables("HTTP_HOST")
end if
 
dim strs:strs=split(advlink,"/")(2)
if request.Cookies("advlink") = 0 then
 
 dim asql,ars
 set ars=server.CreateObject("adodb.recordset")
 asql="select * from NwebCn_Ads_effect where ADS_Link = '"&advlink&"'"
 ars.open asql,conn,1,3
 if not ars.eof then
     ars("ipcount") = ars("ipcount") + 1
	 ars.update
	 Response.Cookies("advlink") = ars("Id")
	 conn.execute("insert into NwebCn_Ip (adv_id,ip,addtime) Values("&ars("Id")&",'"&userip&"','"&now&"')")
	 else
	  ars.close
	  asql="select * from NwebCn_Ads_effect where ADS_Link = '"&strs&"'"
	  ars.open asql,conn,1,3
	  if not ars.eof then
	    ars("ipcount") = ars("ipcount") + 1
	    ars.update
	    Response.Cookies("advlink") = ars("Id")
		conn.execute("insert into NwebCn_Ip (adv_id,ip,addtime) Values("&ars("Id")&",'"&userip&"','"&now&"')")
	  else
	    Response.Cookies("advlink") = 0
	  end if
 end if
 ars.close
 set rs=nothing
end if
%>
