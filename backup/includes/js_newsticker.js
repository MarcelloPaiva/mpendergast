function PopNews(Id) {
  var winl = (screen.width - 600) / 2;
  var wint = (screen.height - 600) / 2;
  window.open("newsDetail.asp?Id=" + Id,"News","width=600,height=600,status=yes,scrollbars=yes,resizable=yes,titlebar=0,left="+winl+",top="+wint);
}
/* USE WORDWRAP AND MAXIMIZE THE WINDOW TO SEE THIS FILE
========================================
 V-NewsTicker v2.1
 License : Freeware (Enjoy it!)
 (c)2003 VASIL DINKOV- PLOVDIV, BULGARIA
========================================
 For IE4+, NS4+ & Opera5+
========================================
 Get the NewsTicker script at:
 http://www.smartmenus.org/other.php
 and don't wait to get the Great SmartMenus script at:
 http://www.smartmenus.org
 LEAVE THESE NOTES PLEASE - delete the comments if you want */

// BUG in Opera:
// If you want to be able to control the body margins
// put the script right after the BODY tag, not in the HEAD!!!

// === 1 === FONT, COLORS, EXTRAS...
v_font='verdana,arial,sans-serif';
v_fontSize='8pt';
v_fontSizeNS4='11px';
v_fontWeight='normal';
v_fontColor='#4A49A8';
v_textDecoration='none';
v_fontColorHover='#996633';//		| won't work
v_textDecorationHover='underline';//	| in Netscape4
v_bgColor='transparent';
// set [='transparent'] for transparent
// set [='url(image_source)'] for image
v_top=10;//	|
v_left=3;//	| defining
v_width=290;//	| the box
v_height=110;//	|
v_paddingTop=0;
v_paddingLeft=10;
v_position='relative';// absolute/relative
v_timeout=5000;//1000 = 1 second
v_slideSpeed=50;
v_slideDirection=0;//0=down-up;1=up-down
v_pauseOnMouseOver=true;
// v2.2+ new below
v_slideStep=1;//pixels
v_textAlign='left';// left/center/right
v_textVAlign='middle';// top/middle/bottom - won't work in Netscape4

// === 2 === THE CONTENT - ['href','text','target']
// Use '' for href to have no link item
/*
v_content=[
['http://www.smartmenus.org/other.php','<img src=strelka.gif align=top width=20 height=11 border=0>Welcome to the V-NewsTicker example page! Presenting the best FREE vertical news scroller ever written.','_blank'],
['','<img src=strelka.gif align=top width=20 height=11 border=0>Featuring: support for the most popular browsers, easy setup, small size, pausing, sliding up or down...','_blank'],
['http://www.smartmenus.org/','<img src=strelka.gif align=top width=20 height=11 border=0>Don\'t wait and also get the most advanced navigation system for your site- the SmartMenus DHTML menu.','_blank']
];
*/
v_content=arrContent;

// THE SERIOUS SCRIPT - PLEASE DO NOT TOUCH
v_ua=navigator.userAgent;
v_nS4=document.layers?1:0;
v_iE=document.all&&!window.innerWidth&&v_ua.indexOf("MSIE")!=-1?1:0;
v_oP=v_ua.indexOf("Opera")!=-1&&document.clear?1:0;
v_oP4=v_ua.indexOf("Opera")!=-1&&!document.clear;
v_kN=v_ua.indexOf("Konqueror")!=-1&&parseFloat(v_ua.substring(v_ua.indexOf("Konqueror/")+10))<3.1?1:0;
v_count=v_content.length;
v_s1=1;
v_cur=1;
v_cl=0;
v_d=v_slideDirection?-1:1;
v_TIM=0;
v_fontSize2=v_nS4&&navigator.platform.toLowerCase().indexOf("win")!=-1?v_fontSizeNS4:v_fontSize;
v_canPause=0;

function v_getOS(a){
  return v_iE?document.all[a].style:v_nS4?document.layers['v_container'].document.layers[a]:document.getElementById(a).style
};

function v_start(){
  var o=v_getOS('v_1');
  var px=v_nS4?0:'px';
  if(v_iE||v_oP){
    if(o.pixelTop==v_paddingTop){
      v_canPause=1;
      if(v_count>1) v_TIM=setTimeout("v_canPause=0;v_slide()",v_timeout);
      return
    };
    o.pixelTop-=v_d;
    if(v_oP&&o.visibility.toLowerCase()!='visible') o.visibility='visible';
    setTimeout("v_start()",v_slideSpeed);
    return
  };
  if(parseInt(o.top)==v_paddingTop){
    v_canPause=1;
    if(v_count>1) v_TIM=setTimeout("v_canPause=0;v_slide()",v_timeout);
    return
  };
  o.top=parseInt(o.top)-v_d+px;
  setTimeout("v_start()",v_slideSpeed)
};

function v_slide(){
  var o=v_getOS('v_'+v_cur);
  var o2=v_getOS('v_'+(v_cur<v_count?v_cur+1:1));
  var px=v_nS4?0:'px';
  if(v_iE||v_oP){
    if(o.pixelTop==v_paddingTop*2-v_height*v_d){
	  if(v_oP) o.visibility='hidden';
      o.pixelTop=v_height*v_d;
      v_cur=v_cur<v_count?v_cur+1:1;
      v_canPause=1;
      v_TIM=setTimeout("v_canPause=0;v_slide()",v_timeout);
      return
	};
    o.pixelTop-=v_d;
    if(v_oP&&o2.visibility.toLowerCase()!='visible') o2.visibility='visible';
    o2.pixelTop-=v_d;
    setTimeout("v_slide()",v_slideSpeed);
    return
  };
  if(parseInt(o.top)==v_paddingTop*2-v_height*v_d){
    o.top=v_height*v_d+px;
    v_cur=v_cur<v_count?v_cur+1:1;
    v_canPause=1;
    v_TIM=setTimeout("v_canPause=0;v_slide()",v_timeout);
    return
  };
  o.top=parseInt(o.top)-v_d+px;
  o2.top=parseInt(o2.top)-v_d+px;
  setTimeout("v_slide()",v_slideSpeed)
};

if(v_nS4||v_iE||v_oP||document.getElementById&&!v_kN&&!v_oP4){
  document.write('<style>.vnewsticker,a.vnewsticker{font-family:'+v_font+';font-size:'+v_fontSize2+';color:'+v_fontColor+';text-decoration:'+v_textDecoration+';font-weight:'+v_fontWeight+'}a.vnewsticker:hover{font-family:'+v_font+';font-size:'+v_fontSize2+';color:'+v_fontColorHover+';text-decoration:'+v_textDecorationHover+'}</style>');
  v_temp='<div '+(v_nS4?"name":"id")+'=v_container style="position:'+v_position+';top:'+v_top+'px;left:'+v_left+'px;width:'+v_width+'px;height:'+v_height+'px;background:'+v_bgColor+';layer-background'+(v_bgColor.indexOf("url(")==0?"-image":"-color")+':'+v_bgColor+';clip:rect(0,'+v_width+','+v_height+',0);overflow:hidden">';
  for(v_i=0;v_i<v_count;v_i++)
    v_temp+='<div '+(v_nS4?"name":"id")+'=v_'+(v_i+1)+' style="position:absolute;top:'+(v_height*v_d)+'px;left:'+v_paddingLeft+'px;width:'+(v_width-v_paddingLeft*2)+'px;height:'+(v_height-v_paddingTop*2)+'px;clip:rect(0,'+(v_width-v_paddingLeft*2)+','+(v_height-v_paddingTop*2)+',0);overflow:hidden'+(v_oP?";visibility:hidden":"")+'">'+(v_content[v_i][0]!=''?'<a href="'+v_content[v_i][0]+'" target="'+v_content[v_i][2]+'" '+(v_pauseOnMouseOver?" onmouseover=\'if(v_canPause&&v_count>1){clearTimeout(v_TIM);v_cl=1}\' onmouseout=\'if(v_canPause&&v_count>1&&v_cl)v_TIM=setTimeout(\"v_canPause=0;v_slide();v_cl=0\","+v_timeout+")\'":"")+'>':'<span'+(v_pauseOnMouseOver?" onmouseover=\'if(v_canPause&&v_count>1){clearTimeout(v_TIM);v_cl=1}\' onmouseout=\'if(v_canPause&&v_count>1&&v_cl)v_TIM=setTimeout(\"v_canPause=0;v_slide();v_cl=0\","+v_timeout+")\'":"")+'>')+v_content[v_i][1]+(v_content[v_i][0]!=''?'</a>':'</span>')+'</div>';v_temp+='</div>';document.write(v_temp);
  setTimeout("v_start()",1000);
  if(v_nS4)onresize=function(){
    location.reload()
  }
}
