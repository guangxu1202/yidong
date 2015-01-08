<!--#include file="include/webtop.asp"-->
<!DOCTYPE HTML>
<html xmlns="http://www.w3.org/1999/xhtml">
 <head>
  <title> <%=t("栏目名称")%> - <%=t("网站名称")%> </title>
  <meta name="Author" content="jam">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <link href="style/theme.css" type="text/css" rel="stylesheet" />
 </head>

<body>
<!--#include file="top.asp"-->
  
  <div class="guide">
	  <div class="warp">
		您的位置：<a href="default.asp">首页</a> &gt;&gt; <a href="range.asp?sortid=13">练习场地</a>
	  </div>
  </div>
  
  <div class="warp">
  	<!--#include file="ileft.asp"-->
	<div class="iright">
		<div class="mainBox productList">
			<div class="mainTop">
				<img src="images/title_zone.gif" />
			</div>
			<div class="mainBody">
				<%=T("newsList$0,新闻列表样式,15,,yyyy-mm-dd,50,y,50,n,y")%>
			<div class="clear"></div>
			</div>
		</div>
	</div>
	<div class="clear"></div>
  </div>
  
  




<!--#include file="footer.asp"-->
</body>
</html>