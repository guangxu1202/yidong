<!--#include file="include/webtop.asp"-->
<!DOCTYPE HTML>
<html xmlns="http://www.w3.org/1999/xhtml">
 <head>
  <title><%=t("新闻标题")%> - <%=t("栏目名称")%> - <%=t("网站名称")%></title>
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
		您的位置：<a href="default.asp">首页</a> &gt;&gt; <a href="mainProlist.asp?sortid=12">主打产品</a>
	  </div>
  </div>
  
  
  <div class="warp">
  	<!--#include file="ileft.asp"-->
	<div class="iright">
		<div class="mainBox productList">
			<div class="mainTop">
				<img src="images/title_main.gif" />
			</div>
			<div class="mainBody mainPro">
			 	 <h3><%=t("产品名称")%></h3>
			 	 <div class="MPdiscription">
					<%=T("产品小图")%>
					<ul>
						<li><span>价格：</span><b><%=T("主打产品价格")%>元</b></li>
						<li><span>车辆：</span><%=T("主打产品车辆")%></li>
						<li><span>机票：</span><%=T("主打产品机票")%></li>
						<li><span>高尔夫：</span><%=T("主打产品高尔夫")%></li>
						<li><span>餐饮：</span><%=T("主打产品餐饮")%></li>
						<li><span>酒店：</span><%=T("主打产品酒店")%></li>
						<li><span>其他：</span><%=T("主打产品其他")%></li>
					</ul>
					<div class="clear"></div>
				 </div>
				 <div class="MPtabs">
					<ul>
						<li class="checked" onClick="changeTab(0)">行程</li>
						<li onClick="changeTab(1)">包含/不包含</li>
						<li onClick="changeTab(2)">相关介绍</li>
						<li onClick="changeTab(3)">注意事项</li>
					</ul>
					<div class="clear"></div>
				 </div>
				 <div class="mpBody">
					 <%=T("产品内容")%>
				 </div>
				 <div class="mpBody" style="display:none; padding:20px;">
				 <%=T("主打产品包含")%>
				 </div>
				 <div class="mpBody" style="display:none; padding:20px;">
				 <%=T("主打产品相关介绍")%>
				 </div>
				 <div class="mpBody" style="display:none; padding:20px;">
				 <%=T("主打产品注意事项")%>
				 </div>
			 
			 
				<div class="clear"></div>
			</div>
		</div>
	</div>
	<div class="clear"></div>
  </div>
  
  
  


<!--#include file="footer.asp"-->
</body>
</html>