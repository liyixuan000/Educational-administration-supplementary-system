// JavaScript Document
var $=function(tagName)
{
	return document.getElementById(tagName);
}
//������Լ��------------------------------------------------------------------------------0
var flag0=[0,0,0];
function check_MesTitle(s)
{
	if(s=="")
	{
		$("span_MesTitle").innerHTML="<img src='../admin/images/yesno.gif'>&nbsp;���������";
		flag0[0]=0;
	}
	else
	{
		$("span_MesTitle").innerHTML="<img src='../admin/images/yesok.gif'>";
		flag0[0]=1;
	}
	check_data0();
}
function check_LinkName(s)
{
	if(s==0)
	{
		$("span_LinkName").innerHTML="<img src='../admin/images/yesno.gif'>&nbsp;����������";
		flag0[1]=0;
	}
	else
	{
		$("span_LinkName").innerHTML="<img src='../admin/images/yesok.gif'>";
		flag0[1]=1;
	}
	check_data0();
}
function check_Content(s)
{
	if(s=="")
	{
		$("span_Content").innerHTML="<img src='../admin/images/yesno.gif'>&nbsp;����������";
		flag0[2]=0;
	}
	else
	{
		$("span_Content").innerHTML="<img src='../admin/images/yesok.gif'>";
		flag0[2]=1;
	}
	check_data0();
}
function check_data0()
{
	if(flag0[0]==1 && flag0[1]==1 && flag0[2]==1){
		$("Btn_OK").disabled=false;
	}
	else{
		$("Btn_OK").disabled=true;
	}
}