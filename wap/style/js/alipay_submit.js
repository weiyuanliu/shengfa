function jisuan()
  {
  if(document.On_Order.On_Numbers1.value=='体验装1支(1)'){
  	var nm1=1*document.On_Order.price1.value;
  }
  else if(document.On_Order.On_Numbers1.value=='体验装1支(2)'){
  	var nm1=2*document.On_Order.price1.value;
  }
  else if(document.On_Order.On_Numbers1.value=='体验装1支(3)'){
  	var nm1=3*document.On_Order.price1.value;
  }
  else if(document.On_Order.On_Numbers1.value=='体验装1支(4)'){
  	var nm1=4*document.On_Order.price1.value;
  }
  else if(document.On_Order.On_Numbers1.value=='体验装1支(5)'){
  	var nm1=5*document.On_Order.price1.value;
  }
  else if(document.On_Order.On_Numbers1.value=='体验装1支(6)'){
  	var nm1=6*document.On_Order.price1.value;
  }
  else{
  	var nm1=0;
  }
  if(document.On_Order.On_Numbers2.value=='改善装2支(1)'){
  	var nm2=1*document.On_Order.price2.value;
  }
  else if(document.On_Order.On_Numbers2.value=='改善装2支(2)'){
  	var nm2=2*document.On_Order.price2.value;
  }
  else if(document.On_Order.On_Numbers2.value=='改善装2支(3)'){
  	var nm2=3*document.On_Order.price2.value;
  }
  else if(document.On_Order.On_Numbers2.value=='改善装2支(4)'){
  	var nm2=4*document.On_Order.price2.value;
  }
  else if(document.On_Order.On_Numbers2.value=='改善装2支(5)'){
  	var nm2=5*document.On_Order.price2.value;
  }
  else if(document.On_Order.On_Numbers2.value=='改善装2支(6)'){
  	var nm2=6*document.On_Order.price2.value;
  }
  else{
  	var nm2=0;
  }										
  document.On_Order.tprice.value=nm1+nm2;
  }