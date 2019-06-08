window.onscroll = function topbarFollow(){  
  var objs = document.getElementsByTagName("div");  
  var scrollt = document.body.scrollTop || document.documentElement.scrollTop;

  for(var i=0;i<objs.length;i++){  

    //topbar
    if(objs[i].id == "topbar"){
      if(scrollt <= objs[i].style.height){
        objs[i].style.top = 0;
        objs[i].style.opacity = 1;
      }
      else{
        objs[i].style.top = scrollt;
        objs[i].style.opacity = 0.9;
      }
    }

  }

  return;  
}  

function godiv(id){  
  var objs = document.getElementsByTagName("div");  

  for(var i=0;i<objs.length;i++){  

    if(objs[i].id == "sdiv" + id){
        window.scrollTo({ 
            top: objs[i].getBoundingClientRect().top, 
            behavior: "smooth" 
        });        
        //document.body.scrollTop = objs[i].getBoundingClientRect().top;
        //document.documentElement.scrollTop = objs[i].getBoundingClientRect().top;
        break;
    }

  }

  return;
}  