/*按比例生成缩略图*/
function DrawImage(ImgD,W,H){ 
  var flag=false; 
  var image=new Image(); 
  image.src=ImgD.src; 
  if(image.width>0 && image.height>0){ 
    flag=true; 
    if(image.width/image.height>= W/H){ 
      if(image.width>W){
        ImgD.width=W; 
        ImgD.height=(image.height*H)/image.width; 
      }
	  else{ 
        ImgD.width=image.width;
        ImgD.height=image.height; 
      } 
      ImgD.alt= ""; 
    } 
    else{ 
      if(image.height>H){
        ImgD.height=H; 
        ImgD.width=(image.width*W)/image.height; 
      }
	  else{ 
        ImgD.width=image.width;
        ImgD.height=image.height; 
      } 
      ImgD.alt=""; 
    } 
  }
}


function createxmlhttp()
{
	var xmlHttp;
	if(window.ActiveXObject)
	{
		try{
			xmlHttp=new ActiveXObject("Msxml2.XMLHTTP");
			} catch(e)
				{
					xmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
				}
	}
	else if(window.XMLHttpRequest)
	{
		xmlHttp=new XMLHttpRequest();
		if(xmlHttp.overrideMimeType)
		{
			xmlHttp.overrideMimeType("text/xml");	
		}
	}
	if(!xmlHttp)
	{
		alert("无法创建xmlHttp对象.");	
		return null;
	}
	return xmlHttp;
}