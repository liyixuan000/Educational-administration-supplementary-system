<!--
function SelectSingle(id)
{
	var L
	L = document.getElementsByName("id").length;
	if(L>1)
	{
	    for(var i=0;i<L;i++)
	    {
		    if(document.all.id[i].value==id)
		    {
			    document.all.id[i].checked=true;
		    }
		    else
		    {
			    document.all.id[i].checked=false;
	        }
	    }
	} 
	else
	{
		document.all.id.checked=true;
	}
}

function SelectAll(yn)
{
    var L
	L = document.getElementsByName("id").length;
	if(L>1)
	{
	    for(var i=0;i<L;i++)
	    {
	        if(yn=="y")
		    {
		        document.all.id[i].checked=true;
		    }
		    else
		    {
		        document.all.id[i].checked=false;
		    }
	    }
	}
	else if(L==1)
	{
	    if(yn="y")
		{
	        document.all.id.checked=true;
		}
		else
		{
		    document.all.id.checked=false;
		}
	}
}
function checkSelectAll(form,page,lag)
{
    var L
	L = document.getElementsByName("id").length;
	if(L>1)
    {
         for(var i=0;i<document.getElementsByName("id").length;i++)
	     {
	         if(document.all.id[i].checked==true)
		         {
                     switch(page)
	                 {
						case "del_class":
						if(! confirm("���༰���������Ѷ������ɾ����ȷʵҪɾ����"))return false;
		                form.action = "activityclass_del.asp?lag="+lag;
						break;
	                    case "del":
		                if(! confirm("ȷʵҪɾ��ѡ�еļ�¼��?"))return false;
		                form.action = "photo_del.asp?lag="+lag;
						break; 
						case "hidden":
						if(! confirm("ȷʵҪ��ѡ�еļ�¼��Ϊδ�����?"))return false;
						form.action = "sethidden_online.asp?action=hidden"+"&lag="+lag;
						break; 
						case "online":
						if(! confirm("ȷʵҪ��ѡ�е���Ϊ�����?"))return false;
						form.action = "sethidden_online.asp?action=online"+"&lag="+lag;
						break;
	                 }
			         form.submit();
				     return true;
		         }
	     }
		 return false;
    }
	
	var L
	L = document.getElementsByName("id").length;
	if(L==1)
	{
	    if(document.all.id.checked==true)
		{
		    switch(page)
			{
				case "del_class":
				if(! confirm("���༰���������Ѷ������ɾ����ȷʵҪɾ����"))return false;
		        form.action = "activityclass_del.asp?lag="+lag;
				break;
			    case "del":
				if(!confirm("ȷʵҪɾ��ѡ���ļ�¼��"))return false;
				form.action = "photo_del.asp?lag="+lag;
				break;
				case "hidden":
				if(! confirm("ȷʵҪ��ѡ�еļ�¼��Ϊδ�����?"))return false;
				form.action = "sethidden_online.asp?action=hidden"+"&lag="+lag;
				break; 
				case "online":
				if(! confirm("ȷʵҪ��ѡ�е���Ϊ�����?"))return false;
				form.action = "sethidden_online.asp?action=online"+"&lag="+lag;
				break;
			}
			form.submit();
			return true;
		}
		return false;
	}
}
function SonsetOrders(form,id,num,lag)
{
	form.action="SonsetOrders.asp?action=SonsetOrders&id="+id+"&num="+num+"&lag="+lag;
	form.submit();
}
function setOrders(form,id,num,lag)
{
	form.action="setOrders.asp?action=setOrders&id="+id+"&num="+num+"&lag="+lag;
	form.submit();
}

function setClassOrders(form,class_id,num,lag)
{
	form.action="setClassOrders.asp?action=setClassOrders&class_id="+class_id+"&num="+num+"&lag="+lag;
	form.submit();
}

function checkDel()
{
    if(! confirm("ȷʵҪɾ��ѡ���ļ�¼��"))return false;
	return true;
}

function checkDelClass()
{
    if(! confirm("���༰���������Ѷ������ɾ����ȷʵҪɾ����"))return false;
	return true;
}

function nwin(id)
{
    win = open('showNews.asp?id='+id,'','left=50,top=10,width=700,height=500,scrollbars=yes')
}
//-->