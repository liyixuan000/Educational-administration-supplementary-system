function document.oncontextmenu(){event.returnValue=false;} //ÆÁ±ÎÊó±êÓÒ¼ü          
function document.onkeydown()         
{         
 if ((window.event.altKey)&&((window.event.keyCode==37)||(window.event.keyCode==39)))  //ÆÁ±ÎAlt+·½Ïò¼ü¡û ¡ú         
 {
  mevent.returnValue=false;         
 }         
                            
 if (     //(event.keyCode==8) ||  ÆÁ±ÎÍË¸ñÉ¾³ı¼ü         
    (event.keyCode==116)|| //ÆÁ±ÎF5 Ë¢ĞÂ¼ü         
    (event.ctrlKey &&event.keyCode==82))
    {        
     event.keyCode=0;         
     event.returnValue=false;         
    }         
}         