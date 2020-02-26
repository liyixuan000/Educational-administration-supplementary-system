//-----------------------------------------------
//showModalDialog Windows
//-----------------------------------------------

function showModal(url,width,height)
{
	try
	{
		window.showModalDialog(url,window,'dialogHeight:'+height+' px; dialogWidth:'+width+' px; dialogTop: px; dialogLeft: px; edge: Raised; center: Yes; help: No; resizable: No; status: No; scroll:No');
	}
	catch(e){}	
}

//-----------------------------------------------
//Open Window
//-----------------------------------------------

function showOpen(url,width,height)
{
	if(!width)
	{
		width = window.screen.availWidth;
	}

	if(!height)
	{
		height = window.screen.availHeight;
	}

	var win=window.open(url,"","toolbar=no,resizable=no,scrollbars=yes,location=no,directories=no,status=no,menubar=no");
	win.moveTo((window.screen.availWidth-width)/2,(window.screen.availHeight-height)/2);
	win.resizeTo(width,height);
}

//-----------------------------------------------
//Open Print
//-----------------------------------------------

function showPrint(url)
{
	var width	= window.screen.availWidth;
	var height	= window.screen.availHeight;

	win=window.open(url,"","toolbar=no,resizable=yes,scrollbars=yes,location=no,directories=no,status=no,menubar=no")
	win.moveTo(0,0)
	win.resizeTo(width,height)
}

//-----------------------------------------------
//Close Window
//-----------------------------------------------

function frm_close()
{
	top.close();
}

function frm_close2()
{
	if (confirm('确定要关闭吗？')==true)
	{top.close();}
}

//-----------------------------------------------
//isnumber
//-----------------------------------------------

function datagrid_isnumber(obj)
{
	number=obj.value;
	if (isNaN(number) || number=="")
	{
		obj.value="";
	}
	else
	{
		if (number>9999999999)
		{
			obj.value="";
		}
	}
}

//-----------------------------------------------
//日期验证
//-----------------------------------------------

function checkDate(strDate)
{
	var result = strDate.match(/((^((1[8-9]\d{2})|([2-9]\d{3}))(-)(10|12|0?[13578])(-)(3[01]|[12][0-9]|0?[1-9])$)|(^((1[8-9]\d{2})|([2-9]\d{3}))(-)(11|0?[469])(-)(30|[12][0-9]|0?[1-9])$)|(^((1[8-9]\d{2})|([2-9]\d{3}))(-)(0?2)(-)(2[0-8]|1[0-9]|0?[1-9])$)|(^([2468][048]00)(-)(0?2)(-)(29)$)|(^([3579][26]00)(-)(0?2)(-)(29)$)|(^([1][89][0][48])(-)(0?2)(-)(29)$)|(^([2-9][0-9][0][48])(-)(0?2)(-)(29)$)|(^([1][89][2468][048])(-)(0?2)(-)(29)$)|(^([2-9][0-9][2468][048])(-)(0?2)(-)(29)$)|(^([1][89][13579][26])(-)(0?2)(-)(29)$)|(^([2-9][0-9][13579][26])(-)(0?2)(-)(29)$))/);
	if(result==null)
	{
		return false;
	}
	return true;
}
//-----------------------------------------------

function killErrors()  
{  
return true;  
}    
window.onerror = killErrors;