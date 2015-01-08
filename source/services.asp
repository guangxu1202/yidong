<!--#include file="include/webtop.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=t("网站名称")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="main">
  <div class="header">
    <div class="logo">
      <h2><%=t("网站名称")%></h2>
      <div class="text"><%=T("网站说明")%></div>
    </div>
    <div class="search">
      <form id="form1" name="form1" method="get" action="newssearch.asp">
        <input name="keyword" type="text"  class="keywords"/>
        <input name="" type="image" src="images/search.gif" onclick="this.submit();" />
      </form>
    </div>
    <div class="clr"></div>
    <div class="menu">
       <%=T("栏目$一级导航栏目")%>
    </div>
    <div class="clr"></div>
      </div>
  <div class="clr"></div>
  <div class="body">
  <h4>About Us</h4>
    <div class="about_body">
      <h2> <%=T("单页名称")%></h2>
      <div>
          <%=T("新闻列表$8,服务列表样式,3,,,0,y,80,n,n")%>
      </div>
    </div>
    <div class="right_body">
      <h2> 最新新闻</h2>
		<%=T("新闻列表$2,首页新闻列表样式,3,,yyyy.mm.dd,0,y,60,n,n")%>
    </div>
    <div class="clr"></div>
  </div>
</div>
   <div class="footer"><p><%=T("网站版权")%></p>
    <p><%=T("html$页脚连接")%></p></div>
 <div class="clr"></div>
</body>
</html>