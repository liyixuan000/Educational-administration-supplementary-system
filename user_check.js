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
	                    case "del":
		                if(! confirm("确实要删除选中的记录吗?"))return false;
		                form.action = "user_del.asp?lag="+lag;
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
			    case "del":
				if(!confirm("确实要删除选定的记录吗？"))return false;
				form.action = "user_del.asp?lag="+lag;
				break;
			}
			form.submit();
			return true;
		}
		return false;
	}
}

function checkDel()
{
    if(! confirm("确实要删除选定的记录吗？"))return false;
	return true;
}


//-->