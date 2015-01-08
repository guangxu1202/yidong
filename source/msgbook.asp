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
		您的位置：<a href="default.asp">首页</a> &gt;&gt; <a href="#"><%=T("栏目名称")%></a>
	  </div>
  </div>
 <div class="warp">
  	<!--#include file="ileft.asp"-->
	<div class="iright">
		<div class="mainBox productList">

			<div class="MPtabs">
				<ul>
					<li class="checked"><%=T("栏目名称")%></li>
				</ul>
				<div class="clear"></div>
			 </div>
			<div class="mainBody">
				
				<%=T("留言列表$0,留言列表样式,5,yyyy-mm-dd,50,y")%>
				<div class="clear"></div>
				<div class="MPtabs">
					<ul>
						<li class="checked">发表留言</li>
					</ul>
					<div class="clear"></div>
				 </div>
     			<%=T("发表留言$留言样式")%>
				<div class="clear"></div>
			</div>
		</div>
	</div>
	<div class="clear"></div>
  </div>



    


<!--#include file="footer.asp"-->
</body>
</html>