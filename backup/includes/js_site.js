function PopNews(Id) {
  var winl = (screen.width - 500) / 2;
  var wint = (screen.height - 600) / 2;
  window.open("news.asp?Id=" + Id,"News","width=500,height=600,status=yes,scrollbars=yes,resizable=yes,titlebar=0,left="+winl+",top="+wint);
}