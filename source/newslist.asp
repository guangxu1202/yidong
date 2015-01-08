<!--#include file="include/webtop.asp"-->
<!DOCTYPE html>
<html>
 <head>
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="description" content="<%=t("wd")%>" />
    <meta name="keywords" content="<%=t("wk")%>" />
  <title> <%=t("栏目名称")%> - <%=t("网站名称")%> </title>
      <!-- css -->
    <link type="text/css" rel="stylesheet" href="style/util.css" media="screen" />
    <link type="text/css" rel="stylesheet" href="style/reset_html5.css" media="screen" />
    <link type="text/css" rel="stylesheet" href="style/index.css" media="screen" />
    
    <!-- js -->
    <script type="text/javascript" src="js/modernizr.2.7.1.js"></script>
    <script type="text/javascript" src="js/TweenMax.min.js"></script>
	<script type="text/javascript" src="js/jquery-1.10.2.min.js"></script>
	<script type="text/javascript" src="js/jquery.flexslider.js"></script>
	<script type="text/javascript" src="js/index.js"></script>
	<script type="text/javascript" src="js/slide.js"></script>
	<script type="text/javascript" src="js/util.js"></script>
 </head>

<body>
<!--#include file="top.asp"-->

<div class="main <%=banner%>">
	<div class="icontent">
		<div class="ic_bar">
			<div class="listhead"><img src="images/<%=ico%>"><%=getSortName(getsortTID(getsortID()))%></div>
			<div class="listbody">
				<%=T("子栏目")%>
			</div>
		</div>
		<div class="ic_main">
			<div class="im_title"><%=t("栏目名称")%></div>
			<div class="im_content">
            	<%=T("newsList$0,新闻列表样式,15,,yyyy-mm-dd,50,y,50,n,y")%>
			</div>
		</div>
	</div>
</div>

  <!--
  <div class="guide">
	  <div class="warp">
		您的位置：<a href="default.asp">首页</a> &gt;&gt; <a href="#">新闻</a>
	  </div>
  </div>
  
  <div class="warp">

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
  
  -->




<!--#include file="footer.asp"-->
</body>
</html>