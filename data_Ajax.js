var http_requestcity=false;
  function send_cityrequest(url)
  {//��ʼ����ָ������������������ĺ���
    http_requestcity=false;
        //��ʼ��ʼ��XMLHttpRequest����
        if(window.XMLHttpRequest)
		{//Mozilla�����
         http_requestcity=new XMLHttpRequest();
           if(http_requestcity.overrideMimeType)
		   {//����MIME���
            http_requestcity.overrideMimeType("text/xml;charset=gb2312");
           }
        }
        else if(window.ActiveXObject)
		{//IE�����
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
        if(!http_requestcity){//�쳣����������ʵ��ʧ��
         window.alert("����XMLHttp����ʧ�ܣ�");
         return false;
        }
        //ȷ����������ʽ��URL�����Ƿ�ͬ��ִ���¶δ���
    http_requestcity.open("post",url,false);
        http_requestcity.send();
  }
    function send_cityrequest2(url)
  {//��ʼ����ָ������������������ĺ���
    http_requestcity2=false;
        //��ʼ��ʼ��XMLHttpRequest����
        if(window.XMLHttpRequest)
		{//Mozilla�����
         http_requestcity2=new XMLHttpRequest();
           if(http_requestcity2.overrideMimeType)
		   {//����MIME���
            http_requestcity2.overrideMimeType("text/xml;charset=gb2312");
           }
        }
        else if(window.ActiveXObject)
		{//IE�����
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
        if(!http_requestcity2){//�쳣����������ʵ��ʧ��
         window.alert("����XMLHttp����ʧ�ܣ�");
         return false;
        }
        //ȷ����������ʽ��URL�����Ƿ�ͬ��ִ���¶δ���
    http_requestcity2.open("post",url,false);
        http_requestcity2.send();
  }
  //��������Ϣ�ĺ���
  function processcityrequest(class_id,selected){
   send_cityrequest('data_Ajax.asp?class_id='+class_id);
   if(http_requestcity.readyState==4){//�ж϶���״̬
     if(http_requestcity.status==200){//��Ϣ�ѳɹ����أ���ʼ������Ϣ
          processstr=http_requestcity.responseText;
		  changecity(processstr,selected);//�˺���Ϊ��������б�
         }
         else{//ҳ�治����
          alert("���������ҳ�治������");
         }
   }
  }
//��������б���
function changecity(cityListStr,selected)
{
	var city;
document.getElementById("pclass_id").length=0;//���optionѡ��
for(i=0;i<cityListStr.split("|").length-1;i++)
  {
	city = cityListStr.split("|")[i].split(",")
	document.getElementById("pclass_id").options[i]=new Option(city[1],city[0])
	
  }
  if(selected>0)document.getElementById("pclass_id").value=selected
}
  //��������Ϣ�ĺ���
  function processcityrequest2(class_id,selected){
   send_cityrequest2('data_Ajax2.asp?class_id='+class_id);
   if(http_requestcity2.readyState==4){//�ж϶���״̬
     if(http_requestcity2.status==200){//��Ϣ�ѳɹ����أ���ʼ������Ϣ
          processstr2=http_requestcity2.responseText;
		  changecity2(processstr2,selected);//�˺���Ϊ��������б�
         }
         else{//ҳ�治����
          alert("���������ҳ�治������");
         }
   }
  }
//��������б���
function changecity2(cityListStr,selected)
{
	var city;
document.getElementById("pid").length=0;//���optionѡ��
for(i=0;i<cityListStr.split("|").length-1;i++)
  {
	city = cityListStr.split("|")[i].split(",")
	document.getElementById("pid").options[i]=new Option(city[1],city[0])
	
  }
  if(selected>0)document.getElementById("pid").value=selected
}
