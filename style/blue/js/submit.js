function JHshNumberText()
{
if ( !(((window.event.keyCode >= 48) && (window.event.keyCode <= 57)) 
|| (window.event.keyCode == 13) || (window.event.keyCode == 46) 
|| (window.event.keyCode == 45)))
{
window.event.keyCode = 0 ;
}
} 

function checkMail(str){var pattern = /^([a-zA-Z0-9_-_.])+@([a-zA-Z0-9_-])+(\.[a-zA-Z0-9_-])+/;if(pattern.test(str)) return true; else return false}
function isInt(str){var pattern = /^\d{0,}$/;if(pattern.test(str)) return true; else return false;}
function isNum(str){var pattern = /^[\-0-9\.]{0,}$/;if(pattern.test(str)) return true; else return false;}
function trim(str){var pattern = /(^\s+)$/;str = str.replace(pattern, "");var pattern = /(\s+)$/;str = str.replace(pattern, "");return str;}
function VerifyTel(StrTel){
	var reg =/^\d{11}$/ ; 
	if (reg.test(StrTel))
	{
		return true;
	}
	else
	{
		return false;
	}
  }
  
function Verify(frm) {
	if (trim(frm.Name.value) == "")
	 {
	  alert("请填写收货人姓名!");
	  frm.Name.focus();
	  return false;
	 }
	 if (VerifyTel(trim(frm.Mobile.value)) ==false)
	 {
	  alert("请正确填写手机号码！");
	  frm.Mobile.focus();
	  return false;
	 }
	 
	 if (trim(frm.ProductNum1.value) ==0 && trim(frm.ProductNum2.value) ==0)
	 {
	  alert("商品数量不能为0，请选择商品数量！");
	  frm.Address.focus();
	  return false;
	 }
	 
	 
	if (trim(frm.Address.value) == "")
	 {
	  alert("请填写真实地址！");
	  frm.Address.focus();
	  return false;
	 }
	 if (checkMail(frm.Email.value)==false && trim(frm.Email.value)!="")
	 {
	  alert("接收订单邮箱可为空，如填写必须正确！");
	  frm.Email.focus();
	  return false;
	 }
	 
	if (trim(frm.PostCode.value) == "")
	 {
	  alert("请填写邮政编码！");
	  frm.PostCode.focus();
	  return false;
	 }
}


								
function jisuanpay()
{
  var HuiKuan = document.order1.HuiKuan.value

  if(document.order1.HuiKuan.value=='货到付款'){
  
  document.order1.ProductName1.value='体验装1支(398元/盒)';									
  document.order1.price1.value=398;
  document.order1.ProductName2.value='改善装2支(598元/盒)';									
  document.order1.price2.value=598;
  }else{
  document.order1.ProductName1.value='体验装1支(378元/盒)';									
  document.order1.price1.value=378;
  document.order1.ProductName2.value='改善装2支(578元/盒)';									
  document.order1.price2.value=578;  
  
  }
  if(document.order1.Numbers1.value=='体验装1支(1)'){
  	var nm1=1*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(2)'){
  	var nm1=2*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(3)'){
  	var nm1=3*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(4)'){
  	var nm1=4*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(5)'){
  	var nm1=5*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(6)'){
  	var nm1=6*document.order1.price1.value;
  }
  else{
  	var nm1=0;
  }
  if(document.order1.Numbers2.value=='改善装2支(1)'){
  	var nm2=1*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(2)'){
  	var nm2=2*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(3)'){
  	var nm2=3*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(4)'){
  	var nm2=4*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(5)'){
  	var nm2=5*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(6)'){
  	var nm2=6*document.order1.price2.value;
  }
  else{
  	var nm2=0;
  }
document.order1.tprice.value=nm1+nm2;
} 


function jisuan()
  {
  if(document.order1.Numbers1.value=='体验装1支(1)'){
  	var nm1=1*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(2)'){
  	var nm1=2*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(3)'){
  	var nm1=3*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(4)'){
  	var nm1=4*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(5)'){
  	var nm1=5*document.order1.price1.value;
  }
  else if(document.order1.Numbers1.value=='体验装1支(6)'){
  	var nm1=6*document.order1.price1.value;
  }
  else{
  	var nm1=0;
  }
  if(document.order1.Numbers2.value=='改善装2支(1)'){
  	var nm2=1*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(2)'){
  	var nm2=2*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(3)'){
  	var nm2=3*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(4)'){
  	var nm2=4*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(5)'){
  	var nm2=5*document.order1.price2.value;
  }
  else if(document.order1.Numbers2.value=='改善装2支(6)'){
  	var nm2=6*document.order1.price2.value;
  }
  else{
  	var nm2=0;
  }										
  document.order1.tprice.value=nm1+nm2;
  } 