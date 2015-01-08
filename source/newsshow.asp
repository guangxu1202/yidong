<!--#include file="include/webtop.asp"-->
<!DOCTYPE html>
<html>
 <head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="description" content="<%=T("newsabstract")%>--<%=t("wd")%>" />
    <meta name="keywords" content="<%=T("newskeywords")%>--<%=t("wk")%>" />
  <title><%=t("新闻标题")%> - <%=t("栏目名称")%> - <%=t("网站名称")%></title>
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

		<div class="ic_main whole">
			<div class="im_title"><%=t("新闻标题")%></div>
            <div class="ballTag">
                <span>日期：<%=T("文章更新时间$yyyy-mm-dd")%></span>   <span>作者：<%=T("nauthor")%></span>  <span>来源：<%=T("文章来源")%></span>
            </div>
            <%If T("新闻摘要")<>"" then%>
            <div class="MPdiscription">
                <strong>摘要：</strong><%=T("新闻摘要")%>
            </div>
            <%End if%>
			<div class="im_content">
            	
                <p>
                
                 <%=T("新闻内容")%>
                </p>
			</div>
		</div>
	</div>
</div>




<!--#include file="footer.asp"-->
</body>
</html>