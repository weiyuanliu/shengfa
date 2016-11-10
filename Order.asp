<!--#Include file="Head.asp"-->
<%
Dim Action:Action=request("Action")
%>

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

if Action <> 0 then

Dim advlink,userip,advlinks,lailuy,lailu,hostlailu,lurl,flurl
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
If userip = "" Then
	userip = Request.ServerVariables("REMOTE_ADDR") 
end if

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

Dim kefupost,krfuip
kefuip="112.195.133.10"
if userip = kefuip then
kefupost=trim(SafeRequest("Sh_Name","post"))
Response.Cookies("advlink") = 0
if Instr(kefupost,"w") <> 0 or Instr(kefupost,"W") <> 0 then
	userip = "112.195.133.11"
	advlink = "http://www.hanfangyanyi.com/w"
elseif Instr(kefupost,"z") <> 0 or Instr(kefupost,"Z") <> 0 then
	userip = "112.195.133.12"
	advlink = "http://www.hanfangyanyi.com/z"
elseif Instr(kefupost,"m") <> 0 or Instr(kefupost,"M") <> 0 then
	userip = "112.195.133.13"
	advlink = "http://www.hanfangyanyi.com/m"
elseif Instr(kefupost,"f") <> 0 or Instr(kefupost,"F") <> 0 then
	userip = "112.195.133.14"
	advlink = "http://www.hanfangyanyi.com/f"
elseif Instr(kefupost,"e") <> 0 or Instr(kefupost,"E") <> 0 then
	userip = "112.195.133.15"
	advlink = "http://www.hanfangyanyi.com/e"
elseif Instr(kefupost,"q") <> 0 or Instr(kefupost,"Q") <> 0 then
	userip = "112.195.133.16"
	advlink = "http://www.hanfangyanyi.com/q"
elseif Instr(kefupost,"s") <> 0 or Instr(kefupost,"S") <> 0 then
	userip = "112.195.133.17"
	advlink = "http://www.hanfangyanyi.com/s"
elseif Instr(kefupost,"u") <> 0 or Instr(kefupost,"U") <> 0 then
	userip = "112.195.133.18"
	advlink = "http://www.hanfangyanyi.com/u"
elseif Instr(kefupost,"p") <> 0 or Instr(kefupost,"P") <> 0 then
	userip = "112.195.133.19"
	advlink = "http://www.hanfangyanyi.com/p"
elseif Instr(kefupost,"k") <> 0 or Instr(kefupost,"K") <> 0 then
	userip = "112.195.133.20"
	advlink = "http://www.hanfangyanyi.com/k"
elseif Instr(kefupost,"h") <> 0 or Instr(kefupost,"H") <> 0 then
	userip = "112.195.133.21"
	advlink = "http://www.hanfangyanyi.com/h"
end if
end if

'lailu = Request.ServerVariables("HTTP_REFERER")	

'session("advlink")=lailu

'advlink = session("advlink")

'if advlink = "" then
'	advlink = request.ServerVariables("HTTP_HOST")
'end if

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

end if
%>

     <div style="background:url(style/blue/images/datu.gif) center  no-repeat; width:1420px;height:410px;margin:0 auto;"></div>
     <div style="background:url(style/blue/images/header_05.jpg) center  no-repeat; width:1420px;height:111px;margin:0 auto;"></div>
	</div>
  <div id="main">
    <div class="topad1"><img src="style/blue/images/news_top.jpg" width="988" height="66" /></div>
<SCRIPT src="style/blue/js/submit.js" type="text/javascript"></SCRIPT>
    <div class="html">
    <div class="html1">
     <div class="listrightb">
                      <%if Action="Left" then%>
                        <!--#include file="info/RequestLeft.asp"-->
                      <%else%>
                          <!--#include file="info/zxdg_left.asp"-->
                      <%End if%>
       </div>
     </div>
   <div>
  </div>
  </div>
  </div><!--#Include file="Order_Foot.asp"-->