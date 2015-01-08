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







  <div class="warp">
  	<!--#include file="ileft.asp"-->
	<div class="iright">
		<div class="mainBox productList">

			<div class="mainBody mainPro">
			 	 <h3><%=t("产品名称")%></h3>

				 
				 <div class="mpBody">

                      <p class="grey"><%=T("产品说明$100")%></p>
                      <p align="center">更新时间:<%=T("产品更新时间$yyyy-mm-dd")%> 自定义字段<%=T("diy$price")%></p>
                     <%=T("产品小图")%>
                      <%=T("产品内容")%>

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