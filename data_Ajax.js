var http_requestcity=false;
  function send_cityrequest(url)
  {//初始化，指定处理函数，发送请求的函数
    http_requestcity=false;
        //开始初始化XMLHttpRequest对象
        if(window.XMLHttpRequest)
		{//Mozilla浏览器
         http_requestcity=new XMLHttpRequest();
           if(http_requestcity.overrideMimeType)
		   {//设置MIME类别
            http_requestcity.overrideMimeType("text/xml;charset=gb2312");
           }
        }
        else if(window.ActiveXObject)
		{//IE浏览器
         try
		  {
          http_requestcity=new ActiveXObject("Msxml2.XMLHttp");
          }
		 catch(e)
		 {
          try
		 {
          http_requestcity=new ActiveXobject("Microsoft.XMLHttp");
          }catch(e){}
         }
       }
        if(!http_requestcity){//异常，创建对象实例失败
         window.alert("创建XMLHttp对象失败！");
         return false;
        }
        //确定发送请求方式，URL，及是否同步执行下段代码
    http_requestcity.open("post",url,false);
        http_requestcity.send();
  }
    function send_cityrequest2(url)
  {//初始化，指定处理函数，发送请求的函数
    http_requestcity2=false;
        //开始初始化XMLHttpRequest对象
        if(window.XMLHttpRequest)
		{//Mozilla浏览器
         http_requestcity2=new XMLHttpRequest();
           if(http_requestcity2.overrideMimeType)
		   {//设置MIME类别
            http_requestcity2.overrideMimeType("text/xml;charset=gb2312");
           }
        }
        else if(window.ActiveXObject)
		{//IE浏览器
         try
		  {
          http_requestcity2=new ActiveXObject("Msxml2.XMLHttp");
          }
		 catch(e)
		 {
          try
		 {
          http_requestcity2=new ActiveXobject("Microsoft.XMLHttp");
          }catch(e){}
         }
       }
        if(!http_requestcity2){//异常，创建对象实例失败
         window.alert("创建XMLHttp对象失败！");
         return false;
        }
        //确定发送请求方式，URL，及是否同步执行下段代码
    http_requestcity2.open("post",url,false);
        http_requestcity2.send();
  }
  //处理返回信息的函数
  function processcityrequest(class_id,selected){
   send_cityrequest('data_Ajax.asp?class_id='+class_id);
   if(http_requestcity.readyState==4){//判断对象状态
     if(http_requestcity.status==200){//信息已成功返回，开始处理信息
          processstr=http_requestcity.responseText;
		  changecity(processstr,selected);//此函数为构造城市列表
         }
         else{//页面不正常
          alert("您所请求的页面不正常！");
         }
   }
  }
//构造城市列表函数
function changecity(cityListStr,selected)
{
	var city;
document.getElementById("pclass_id").length=0;//清空option选项
for(i=0;i<cityListStr.split("|").length-1;i++)
  {
	city = cityListStr.split("|")[i].split(",")
	document.getElementById("pclass_id").options[i]=new Option(city[1],city[0])
	
  }
  if(selected>0)document.getElementById("pclass_id").value=selected
}
  //处理返回信息的函数
  function processcityrequest2(class_id,selected){
   send_cityrequest2('data_Ajax2.asp?class_id='+class_id);
   if(http_requestcity2.readyState==4){//判断对象状态
     if(http_requestcity2.status==200){//信息已成功返回，开始处理信息
          processstr2=http_requestcity2.responseText;
		  changecity2(processstr2,selected);//此函数为构造城市列表
         }
         else{//页面不正常
          alert("您所请求的页面不正常！");
         }
   }
  }
//构造城市列表函数
function changecity2(cityListStr,selected)
{
	var city;
document.getElementById("pid").length=0;//清空option选项
for(i=0;i<cityListStr.split("|").length-1;i++)
  {
	city = cityListStr.split("|")[i].split(",")
	document.getElementById("pid").options[i]=new Option(city[1],city[0])
	
  }
  if(selected>0)document.getElementById("pid").value=selected
}
