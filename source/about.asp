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
			<div class="im_title"><%=T("单页名称")%></div>
			<div class="im_content">
            	<%=T("单页内容$0,0,html")%>
                
			</div>
		</div>
	</div>
</div>



 <!--#include file="footer.asp"-->

</body>
</html>