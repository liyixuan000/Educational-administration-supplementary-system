var maxtime;
if(window.name==''){ 
maxtime = 60*60;
}else{
maxtime = window.name;
}
function CountDown(){
if(maxtime>=0){
minutes = Math.floor(maxtime/60);
seconds = Math.floor(maxtime%60);
msg = "ʣ��ʱ�䣺"+minutes+"��"+seconds+"��";
document.all["timer"].innerHTML = msg;
if(maxtime == 15*60) alert('ע�⣬����15���ӿ��Լ�������!');
--maxtime;
window.name = maxtime; 
}
else{
clearInterval(timer);
window.alert("����ʱ���ѵ����Ծ����ύ��");window.location="submit.asp";document.form1.submit();
}
}
timer = setInterval("CountDown()",1000);