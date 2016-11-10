<%
Function FlashNews(num,fontnum,PicWidthStr,PicHeightStr,BGcolor,txtheight,x)
dim FilterSql
Dim RsFilterObj,FlashStr,ImagesStr,TxtStr,LinkStr
FilterSql = "select top "&num&" * from NwebCn_Flash where ViewFLag=1 order by px desc,id ASC"
Set RsFilterObj = Conn.Execute(FilterSql)
If not RsFilterObj.Eof then
  Dim Temp_Num
  Temp_Num = 0
  Do While Not RsFilterObj.Eof
  Temp_Num = Temp_Num + 1
  RsFilterObj.MoveNext
  Loop
  RsFilterObj.MoveFirst
  If Temp_Num <=1 then
  Set RsFilterObj = Nothing
  FlashNews = "至少需要两条幻灯新闻才能正确显示幻灯效果"
  Set RsFilterObj = Nothing
  Exit Function 
  End If
  do while Not RsFilterObj.Eof
   if ImagesStr = "" then
     ImagesStr =Or2(RsFilterObj("BigPic"))
     TxtStr = RsFilterObj("Title")
     LinkStr = RsFilterObj("Url")&""
   else
     ImagesStr = ImagesStr &"|" &Or2(RsFilterObj("BigPic"))
     TxtStr = TxtStr&"|"&RsFilterObj("Title")
     LinkStr = LinkStr&"|"&RsFilterObj("Url")&""
   end if
  RsFilterObj.MoveNext
  loop
FlashStr="<script type=""text/javascript"">"& Chr(13)
FlashStr=FlashStr&"<!--"& Chr(13)
FlashStr=FlashStr&"var focus_width="&PicWidthStr& Chr(13)   
FlashStr=FlashStr&"var focus_height="&PicHeightStr& Chr(13) 
FlashStr=FlashStr&"var text_height="&txtheight& Chr(13) 
FlashStr=FlashStr&"var swf_height = focus_height+text_height"& Chr(13)
FlashStr=FlashStr&"var pics='"&ImagesStr&"'"&Chr(13)
FlashStr=FlashStr&"var links='"&LinkStr &"'"&Chr(13)
FlashStr=FlashStr&"var texts='"&TxtStr&"'"&Chr(13)
FlashStr=FlashStr&"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000""codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ focus_width +'"" height=""'+ swf_height +'"">');"&Chr(13)
FlashStr=FlashStr&"document.write('<param name=""allowScriptAccess"" value=""sameDomain""><param name=""movie"" value=""Include/flash.swf""><param name=""quality"" value=""high""><param name=""bgcolor"" value="&BGcolor&">');"&Chr(13)
FlashStr=FlashStr&"document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');"&Chr(13)
FlashStr=FlashStr&" document.write('<param name=""FlashVars"" value=""pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'"">');"&Chr(13)
FlashStr=FlashStr&"document.write('<embed src=""Include/flash.swf"" wmode=""opaque"" FlashVars=""pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'"" menu=""false"" bgcolor="&BGcolor&" quality=""high"" width=""'+ focus_width +'"" height=""'+ swf_height +'"" allowScriptAccess=""sameDomain"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />');"&Chr(13)
FlashStr=FlashStr&"document.write('</object>');"&Chr(13)
FlashStr=FlashStr&"//-->"& Chr(13)
FlashStr=FlashStr&"</script>"
  else
    echo "暂时没有幻灯片"
  end if
    RsFilterObj.Close
Set RsFilterObj = Nothing
    FlashNews= FlashStr
End Function

%>