// JavaScript Document
/* USE WORDWRAP AND MAXIMIZE THE WINDOW TO SEE THIS FILE
========================================
 V-NewsTicker v2.2
 License : Freeware (Enjoy it!)
 (c)2003 VASIL DINKOV- PLOVDIV, BULGARIA
========================================
 For IE4+, NS4+, Opera5+, Konqueror3.1+
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
v_height=90;//	|
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

v_content = [


