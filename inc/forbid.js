function document.oncontextmenu(){event.returnValue=false;} //��������Ҽ�          
function document.onkeydown()         
{         
 if ((window.event.altKey)&&((window.event.keyCode==37)||(window.event.keyCode==39)))  //����Alt+������� ��         
 {
  mevent.returnValue=false;         
 }         
                            
 if (     //(event.keyCode==8) ||  �����˸�ɾ����         
    (event.keyCode==116)|| //����F5 ˢ�¼�         
    (event.ctrlKey &&event.keyCode==82))
    {        
     event.keyCode=0;         
     event.returnValue=false;         
    }         
}         