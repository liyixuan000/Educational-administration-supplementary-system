var gdCtrl = new Object();
var goSelectTag = new Array();
var gcGray = "#cccccc";
var gcToggle = "#FFFF00";
var gcred = "#FF0000";
var gcBG = "#F8F9EE";
var gcGreen = "#00FF00"

var gdCurDate = new Date();
var giYear = gdCurDate.getFullYear();
var giMonth = gdCurDate.getMonth()+1;
var giDay = gdCurDate.getDate();

function fPopCalendar(popCtrl, dateCtrl){
  event.cancelBubble=true;
  gdCtrl = dateCtrl;
  fSetYearMon(giYear, giMonth);
  var point = fGetXY(popCtrl);
  with (VicPopCal.style) {
  	left = point.x;
	top  = point.y+popCtrl.offsetHeight+1;
	width = VicPopCal.offsetWidth;
	height = VicPopCal.offsetHeight;
	fToggleTags(point);
	visibility = 'visible';
  }
  VicPopCal.focus();
}

 function fSetDate(iYear, iMonth, iDay){
	var now=new Date();
	var hours = now.getHours();  //�õ�Сʱ
	var minutes = now.getMinutes();  //�õ�����
	var seconds = now.getSeconds();  //�õ���
  gdCtrl.value = iYear+"-"+iMonth+"-"+iDay+" "+hours+":"+minutes+":"+seconds;
  fHideCalendar();
}

function fHideCalendar(){
  VicPopCal.style.visibility = "hidden";
  for (i in goSelectTag)                                        //��������ѭ��ȡֵ��goSelectTagΪ������
  	goSelectTag[i].style.visibility = "visible";
  goSelectTag.length = 0;
}

function fSetSelected(aCell){                                        
  var iOffset = 0;
  var iYear =parseInt(tbSelYear.value);//��̬�ı��ı������
  var iMonth = parseInt(tbSelMonth.value);//��̬�ı��ı������

  aCell.bgColor = gcBG;
  with (aCell.children["cellText"]){
  	var iDay = parseInt(innerText);
  	if (color==gcGray)
		iOffset = (Victor<10)?-1:1;
	iMonth += iOffset;
	if (iMonth<1) {
		iYear--;
		iMonth = 12;
	}else if (iMonth>12){
		iYear++;
		iMonth = 1;
	}
  }
  fSetDate(iYear, iMonth, iDay);
}

function Point(iX, iY){
	this.x = iX;
	this.y = iY;
}

function fBuildCal(iYear, iMonth) {
  var aMonth=new Array();
  for(i=1;i<7;i++)
  	aMonth[i]=new Array(i);

  var dCalDate=new Date(iYear, iMonth-1, 1);
  var iDayOfFirst=dCalDate.getDay();
  var iDaysInMonth=new Date(iYear, iMonth, 0).getDate();
  var iOffsetLast=new Date(iYear, iMonth-1, 0).getDate()-iDayOfFirst+1;
  var iDate = 1;
  var iNext = 1;

  for (d = 0; d < 7; d++)
	aMonth[1][d] = (d<iDayOfFirst)?-(iOffsetLast+d):iDate++;
  for (w = 2; w < 7; w++)
  	for (d = 0; d < 7; d++)
		aMonth[w][d] = (iDate<=iDaysInMonth)?iDate++:-(iNext++);
  return aMonth;
}

function fDrawCal(iYear, iMonth, iCellHeight, iDateTextSize) {
  var WeekDay = new Array("��","һ","��","��","��","��","��");
  var styleTD = " bgcolor='"+gcBG+"' bordercolor='"+gcBG+"' valign='middle' align='center' height='"+iCellHeight+"' style='font-size:9pt "+iDateTextSize+" ����;";

  with (document) {
	write("<tr>");
	for(i=0; i<7; i++)
		write("<td "+styleTD+"color:green'>" + WeekDay[i] + "</td>");
	write("</tr>");

  	for (w = 1; w < 7; w++) {
		write("<tr>");
		for (d = 0; d < 7; d++) {
			write("<td id=calCell "+styleTD+"cursor:hand;' onMouseOver='this.bgColor=gcToggle' onMouseOut='this.bgColor=gcBG' onclick='fSetSelected(this)'>");
			write("<font id=cellText Victor='Liming Weng'> </font>");
			write("</td>")
		}
		write("</tr>");
	}
  }
}

function fUpdateCal(iYear, iMonth) {
  myMonth = fBuildCal(iYear, iMonth);
  var i = 0;
  for (w = 0; w < 6; w++)
	for (d = 0; d < 7; d++)
		with (cellText[(7*w)+d]) {
			Victor = i++;
			if (myMonth[w+1][d]<0) {
				color = gcGray;
				innerText = -myMonth[w+1][d];
			}else{
				color = ((d==0)||(d==6))?"red":"black";
				innerText = myMonth[w+1][d];
			}
		}
}
//�ú�����̬�ı������������еı仯
function fSetYearMon(iYear, iMon){
  tbSelMonth.options[iMon-1].selected = true;
  for (i = 0; i < tbSelYear.length; i++)
	if (tbSelYear.options[i].value == iYear)
		tbSelYear.options[i].selected = true;
  fUpdateCal(iYear, iMon);//���ı���ֵ����fUpdateCal�����Ա����ڱ�������ʾ�仯
}

function fPrevMonth(){
  var iMon = tbSelMonth.value;
  var iYear = tbSelYear.value;

  if (--iMon<1) {
	  iMon = 12;
	  iYear--;
  }

  fSetYearMon(iYear, iMon);
}

function fNextMonth(){
  var iMon = tbSelMonth.value;
  var iYear = tbSelYear.value;

  if (++iMon>12) {
	  iMon = 1;
	  iYear++;
  }

  fSetYearMon(iYear, iMon);
}

function fToggleTags(){
  with (document.all.tags("SELECT")){
 	for (i=0; i<length; i++)
 		if ((item(i).Victor!="Won")&&fTagInBound(item(i))){
 			item(i).style.visibility = "hidden";
 			goSelectTag[goSelectTag.length] = item(i);
 		}
  }
}

function fTagInBound(aTag){
  with (VicPopCal.style){
  	var l = parseInt(left);
  	var t = parseInt(top);
  	var r = l+parseInt(width);
  	var b = t+parseInt(height);
	var ptLT = fGetXY(aTag);
	return !((ptLT.x>r)||(ptLT.x+aTag.offsetWidth<l)||(ptLT.y>b)||(ptLT.y+aTag.offsetHeight<t));
  }
}

function fGetXY(aTag){
  var oTmp = aTag;
  var pt = new Point(0,0);
  do {
  	pt.x += oTmp.offsetLeft;
  	pt.y += oTmp.offsetTop;
  	oTmp = oTmp.offsetParent;
  } while(oTmp.tagName!="BODY");
  return pt;
}
function fClearInput()
{
	gdCtrl.value = "";
  	fHideCalendar();
}

var gMonths = new Array("&nbsp;һ��","&nbsp;����","&nbsp;����","&nbsp;����","&nbsp;����","&nbsp;����","&nbsp;����","&nbsp;����","&nbsp;����","&nbsp;ʮ��","ʮһ��","ʮ����");

with (document) {
write("<Div id='VicPopCal' onclick='event.cancelBubble=true' style='POSITION:absolute;visibility:hidden;border:1px ridge;width:10;z-index:100;'>");
write("<table border='0' bgcolor='#F8F9EE'>");
write("<TR>");
write("<td valign='middle' align='center'><input type='button' name='PrevMonth' value='��' style='height:20;width:20' onClick='fPrevMonth()'>");
write("&nbsp;<SELECT name='tbSelYear' onChange='fUpdateCal(tbSelYear.value, tbSelMonth.value)' style='font-color:#000080;width:70;border:1 solid #99CCFF; font-size:9pt; background-color:#F8F9EE' Victor='Won'>");
for(i=1968;i<=2028;i++)
	write("<OPTION value='"+i+"'>"+i+"��</OPTION>");
write("</SELECT>");
write("&nbsp;<select name='tbSelMonth' onChange='fUpdateCal(tbSelYear.value, tbSelMonth.value)'  style='font-color:#000080;width:70;border:0 solid #99CCFF; font-size:9pt; background-color:#F8F9EE' Victor='Won'>");
for (i=0; i<12; i++)
	write("<option value='"+(i+1)+"'>"+gMonths[i]+"</option>");
write("</SELECT>");
write("&nbsp;<input type='button' name='PrevMonth' value='��' style='height:20;width:20' onclick='fNextMonth()'>");
write("</td>");
write("</TR><TR>");
write("<td align='center'>");
write("<DIV style='background-color:blue'><table border='0' cellspacing='0' cellpadding='0' width='100%' bgcolor='blue'><tr><td><table border='0' cellspacing='1' width='100%' cellpadding='1'>");
fDrawCal(giYear, giMonth, 12, 12);
write("</table></td></tr></table></DIV>");
write("</td>");
write("</TR><TR><TD align='center'>");
write("<span style='cursor:hand; font-size=9pt' onclick='fSetDate(giYear,giMonth,giDay)' onMouseOver='this.style.color=gcred' onMouseOut='this.style.color=0'>���죺"+giYear+"-"+giMonth+"-"+giDay+"</span>");
write("<span style='cursor:hand; font-size=9pt' onclick='fClearInput()' onMouseOver='this.style.color=gcGreen' onMouseOut='this.style.color=0'>&nbsp;&nbsp;���</span>");
write("</TD></TR>");
write("</TABLE></Div>");
write("<SCRIPT event=onclick() for=document>fHideCalendar()</SCRIPT>");
}

function arrowtag(namestr,valuestr)
{
	  document.write("<input type='text'  name='"+namestr+"' value='"+valuestr+"' size='20' class='wenbenkuang' style='text-align: center;'>&nbsp;<Img src='/images/datetime.gif' style='cursor:hand;' align='absmiddle' alt='�������������˵�' onclick='fPopCalendar("+namestr+","+namestr+");return false'>");
}