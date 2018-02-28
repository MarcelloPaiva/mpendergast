// JavaScript Document

//  holds the name of the subdirectory that contains all your clock images
//  "" means that the images are in the same directory as the script
//  "images/" means that the images are in the images directory
// var image_dir = "liveclock/";
var image_dir = "images/date/";

// update every how many seconds?
var speed = 60;

// ********** NO NEED TO MODIFY ANYTHING BELOW THIS POINT ************

// preload number images
img0 = new Image(); img0.src = image_dir+"0.gif";
img1 = new Image(); img1.src = image_dir+"1.gif";
img2 = new Image(); img2.src = image_dir+"2.gif";
img3 = new Image(); img3.src = image_dir+"3.gif";
img4 = new Image(); img4.src = image_dir+"4.gif";
img5 = new Image(); img5.src = image_dir+"5.gif";
img6 = new Image(); img6.src = image_dir+"6.gif";
img7 = new Image(); img7.src = image_dir+"7.gif";
img8 = new Image(); img8.src = image_dir+"8.gif";
img9 = new Image(); img9.src = image_dir+"9.gif";

// fixes a Netscape 2 and 3 bug
function getFullYear(d) { // d is a date object
 yr = d.getYear();
 if (yr < 1000)
  yr+=1900;
 return yr;
}
function calcDate() {
 now = new Date();
 d = now.getDay();
 document.dayofweek.src = image_dir+"d" + d + ".gif";
 m = now.getMonth();
 document.month.src = image_dir+"m" + m + ".gif";
 dd = now.getDate();
 dd = (dd >= 10 ? "":"0") + dd;
 document.date1.src = image_dir + dd.substring(0,1) + ".gif";
 document.date2.src = image_dir + dd.substring(1,2) + ".gif";
 y = "" + getFullYear(now);
 document.year1.src = image_dir + y.substring(0,1) + ".gif";
 document.year2.src = image_dir + y.substring(1,2) + ".gif";
 document.year3.src = image_dir + y.substring(2,3) + ".gif";
 document.year4.src = image_dir + y.substring(3,4) + ".gif";
 calcAmPm();
}

function calcAmPm(){
 ampm = new Date();
 ap = ampm.getHours();
 if (ap < 12) {
  document.morn.src = image_dir + "am.gif";
 } else {
  document.morn.src = image_dir + "pm.gif";
 }
 calcTime();
}

speed *= 1000; // time kept in milliseconds not seconds

function calcTime() {
 now = new Date();
 h = now.getHours();
 if (h > 12)
  h = h - 12;
 h = (h<10?"0":"") + h;
 document.hour1.src = image_dir + h.substring(0,1) + ".gif";
 document.hour2.src = image_dir + h.substring(1,2) + ".gif";
 m = now.getMinutes();
 m = (m<10?"0":"") + m;
 document.min1.src = image_dir + m.substring(0,1) + ".gif";
 document.min2.src = image_dir + m.substring(1,2) + ".gif";
 //se = now.getSeconds();
 //se = (se<10?"0":"") + se;
 //document.sec1.src = image_dir + se.substring(0,1) + ".gif";
 //document.sec2.src = image_dir + se.substring(1,2) + ".gif";
 if ((h < 1) && (m < 1) && (se < 1))
  calcDate();
 if ((h == 12) && (m < 1) && (se < 1))
  calcAmPm();
 setTimeout("calcTime()",speed);
}

function doDate() {
	document.write('<img name="dayofweek" src="'+image_dir+'space.gif">'
			+'<img src="'+image_dir+'space.gif" width=3 height=6>'
			+'<img name="month" src="'+image_dir+'space.gif"><img src="'
			+image_dir+'space.gif" width=3 height=6><img name="date1" '
			+'src="'+image_dir+'space.gif"><img name="date2" src="'+image_dir+'space.gif" '
			+'><img name="comma" src="'+image_dir+'cc.gif" border=0 '
			+'><img src="'+image_dir+'space.gif" width=3 height=6>'
			+'<img name="year1" src="'+image_dir+'space.gif"><img name="year2" src="" '
			+'><img name="year3" src="'+image_dir+'space.gif">'
			+'<img name="year4" src="'+image_dir+'space.gif"><img src="'+image_dir
			+'space.gif" width=3 height=6><img src="'+image_dir+'space.gif" width=3 height=6>'
			+'<img name="hour1" src="'+image_dir+'space.gif">'
			+'<img name="hour2" src="'+image_dir+'space.gif"><img name="colon" src="'
			+image_dir+'c.gif" border=0><img name="min1" src="'+image_dir+'space.gif" '
			+'><img name="min2" src="'+image_dir+'space.gif">'
			//+'<img name="colon" src="'+image_dir+'c.gif">'
			//+'<img name="sec1" src="'+image_dir+'space.gif" width=9 height=11><img name="sec2" src="" '
			//+'width=9 height=11><img src="'+image_dir+'space.gif" width=9 height=11>'
			+'<img name="morn" src="'+image_dir+'space.gif">');
			calcDate();
	
	}


