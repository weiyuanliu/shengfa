// JavaScript Document
function ChangEvent(Source,Source_Two,Source_Three,FileName)
{
	var Select_1=document.getElementById(Source);
	var xmlhttp=new createxmlhttp();
	var queryString="ParentID="+escape(Select_1.value);
	xmlhttp.onreadystatechange =function(){BackXmlhttp(xmlhttp,Source_Two,Source_Three);};
	xmlhttp.open("POST",FileName,true);
	xmlhttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	xmlhttp.send(queryString);
	
}

function BackXmlhttp(xmlhttp,Source_Two,Source_Three)
{
	if(xmlhttp.readyState==4)
		{
			if(xmlhttp.status==200)
			{
				
				var text=xmlhttp.responseText;
				text=text.slice(text.indexOf("$")+1,text.lastIndexOf("$"));
				
				if(text.indexOf("Error")==-1)
				{
					text=text.slice(0,text.lastIndexOf("|"));
					if(Source_Three!="Null" && text.indexOf("||")!=-1)
					{
						
						var Select1=document.getElementById(Source_Two);
						var Select2=document.getElementById(Source_Three);
						
						
						var ArraySum,ArrayTwo,ArrayThree,items_value;
						
						Select1.length=0;
						Select2.length=0;
						
						ArraySum=text.split("||")
						ArrayTwo=ArraySum[0].split("|")
						ArrayThree=ArraySum[1].split("|")
						
						for(var i=0;i<ArrayTwo.length;i++)
						{
							items_value=ArrayTwo[i].split(",");
							Select1.options[i]=new Option(items_value[1],items_value[0])
						}
						
						for(var j=0;j<ArrayThree.length;j++)
						{
							items_value=ArrayThree[j].split(",");
							Select2.options[j]=new Option(items_value[1],items_value[0])	
						}
						
					}
					else
					{
						var Select=document.getElementById(Source_Two);
						var ArrayValues=text.split("|")
						var item_values;
						
						Select.length=0;
						for(var i=0;i<ArrayValues.length;i++)
						{
							item_values=ArrayValues[i].split(",");
							Select.options[i]=new Option(item_values[1],item_values[0])
						}
					}
				}
				else
				{
					alert("对不起，出现错！");
					window.location.href="Regionallist.asp";
				}
				
			}
			
		}
}


function Check_AddRegionalValues()
{
	var QY_Names,QY_ShengFen,QY_City,QY_Citys,QY_Type,QY_XingZhi,QY_FanWei
	var QY_Wai,QY_CaoZuo,QY_BeiZu,QY_AddTime,QY_Px
	
	QY_Names=document.getElementById("QY_Names");
	QY_ShengFen=document.getElementById("QY_ShengFen");
	QY_City=document.getElementById("QY_City");
	QY_Citys=document.getElementById("QY_Citys");
	QY_Type=document.getElementById("QY_Type");
	QY_XingZhi=document.getElementById("QY_XingZhi");
	QY_FanWei=document.getElementById("QY_FanWei");
	
	QY_Wai=document.getElementById("QY_Wai");
	QY_CaoZuo=document.getElementById("QY_CaoZuo");
	QY_BeiZu=document.getElementById("QY_BeiZu");
	QY_AddTime=document.getElementById("QY_AddTime");
	QY_Px=document.getElementById("QY_Px");
	
	if(QY_Names.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写名字！");
		QY_Names.focus();
		return false;
	}
	
	if(QY_ShengFen.value=="Null")
	{
		alert("对不起，该值不能为空，请选择！");
		QY_ShengFen.focus();
		return false;
	}
	
	if(QY_City.value=="Null")
	{
		alert("对不起，该值不能为空，请选择！");
		QY_City.focus();
		return false;
	}
	
	if(QY_Citys.value=="Null")
	{
		alert("对不起，该值不能为空，请选择！");
		QY_Citys.focus();
		return false;	
	}
	
	if(QY_Type.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("该值不能为空，请填写！");
		QY_Type.focus();
		return false;
	}
	
	if(QY_XingZhi.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("该值不能为空，请填写！");	
		QY_XingZhi.focus();
		return false;
	}
	
	if(QY_FanWei.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写服务范围！");
		return false;
	}
	
	if(QY_Wai.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写服务范围外信息！");	
		return false;
	}
	
	if(QY_Px.value.replace(/^\s*|\s*$/g,'')=="")
	{
		alert("请填写排序顺序！");
		QY_Px.focus();
		return false;
	}
	else
	{
		if((QY_Px.value).search("^-?\\d+(\\.\\d+)?$")!=0)
		{
			alert("请填写数字！");
			QY_Px.select();
			return false;
		}
	}
	return true;
}

