var lastfind = -1;

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

function searchfor(str){  
  var objs = document.getElementsByTagName("span");  
  var finalmark = 0;

  for(var step=0;step<2;step++){
    for(var i=0;i<objs.length;i++){  

      //测试通过寻找元素内容滑动页面
      if((objs[i].innerText.indexOf(str) > 0) && (lastfind != i)){
          objs[i].style.animationName = "searchlight";
          objs[i].style.animationDuration = 2;
          window.scrollTo({ 
              top: objs[i].getBoundingClientRect().top - document.body.clientHeight / 2, 
              behavior: "smooth" 
          });    
          lastfind = i;
          finalmark = 1;  
          break;
      }

    }

    if(finalmark == 1){break;}
    if(step + 1 == 1){objs = document.getElementsByTagName("td");}
  }

  lastfind = -1;
  return;

}  