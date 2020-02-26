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
msg = "剩余时间："+minutes+"分"+seconds+"秒";
document.all["timer"].innerHTML = msg;
if(maxtime == 15*60) alert('注意，还有15分钟考试即将结束!');
--maxtime;
window.name = maxtime; 
}
else{
clearInterval(timer);
window.alert("考试时间已到，试卷即将提交！");window.location="submit.asp";document.form1.submit();
}
}
timer = setInterval("CountDown()",1000);