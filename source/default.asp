<!--#include file="include/webtop.asp"-->

<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="description" content="<%=t("wd")%>" />
    <meta name="keywords" content="<%=t("wk")%>" />
    <title> <%=t("网站名称")%></title> 
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



<div class="index-content">
	<div class="banner">
		<ul class="slides">
			<li>
				<img src="images/banner01.jpg" width='100%' alt="" /> 
				<dl class="text">
					<dt>云行天下  大益东方</dt>
					<dd class="data">Run the world Big profit orient</dd>
				</dl>
			</li>
			<li>
				<img src="images/banner02.jpg" width='100%' alt="" />
				<dl class="text">
					<dt style="height:90px; line-height:40px;">诚信服务立天下伟业<br/>精益求精塑益东品牌</dt>
					<dd class="data">Good faith service made the world Keep improving plastic yi east brand</dd>
				</dl>
			</li>

			<li>
				<img src="images/banner03.jpg" width='100%' alt="" />
				<dl class="text">
					<dt style="height:90px; line-height:40px;">辽宁汤沟国际温泉<br/>养生旅游度假区</dt>
					<dd class="data">Liaoning ThangGou international hot sprin tourist resort</dd>
				</dl>
			</li>

		</ul>
	</div>


	<div class="module">
		<div class="center index-center">
			<div class="kwicks">
				<div class="kwick  first">
					<div>
						<span class="icon icon00"></span>
						<div class="kwick-nr">
							<a href="about.asp?sortid=42">
							<p><img src="images/small_01.jpg" alt="" /></p>
							<h3>益东科技文字简介</h3>
							</a>	
						</div>
					</div>
				</div>				
				<div class="kwick ">
					<div>
						<span class="icon icon01"></span>
						<div class="kwick-nr">
							<a href="about.asp?sortid=43">
							<p><img src="images/small_02.jpg" alt="" /></p>
							<h3>益东金融文字简介</h3>
							</a>	
						</div>
					</div>
				</div>				
				<div class="kwick ">
					<div>
						<span class="icon icon02"></span>
						<div class="kwick-nr">
							<a href="about.asp?sortid=44">
							<p><img src="images/small_03.jpg" alt="" /></p>
							<h3>益东矿业文字简介</h3>
							</a>	
						</div>
					</div>
				</div>				
				<div class="kwick ">
					<div>
						<span class="icon icon03"></span>
						<div class="kwick-nr">
							<a href="about.asp?sortid=45">
							<p><img src="images/small_04.jpg" alt="" /></p>
							<h3>益东旅游文字简介</h3>
							</a>	
						</div>
					</div>
				</div>				
				<div class="kwick ">
					<div>
						<span class="icon icon04"></span>
						<div class="kwick-nr">
							<a href="about.asp?sortid=46">
							<p><img src="images/small_05.jpg" alt="" /></p>
							<h3>益东地产文字简介</h3>
							</a>	
						</div>
					</div>
				</div>				
			</div>
		</div>

	</div>
</div>




<div class="index_footer">
	<div class="iwrap">
		<div class="index_news">
			<div class="in_title">News</div>
			<div class="newsslide">
            	<%=T("newsList$29,首页新闻列表样式,10,,yyyy-mm-dd,50,y,50,n,n")%>
			</div>
			<script>
				var news = $('#news-wrap');
				var child = news.children();
				var len = child.length;
				var page = Math.ceil(len/2);
				for(var i =0;i<page;i++){
					child.slice(i*2,i*2+2).wrapAll('<li></li>');
				}
			</script>
		</div>
		<div class="index_info">
			<div class="atten">
				<span>关注益东：</span>
			</div>
			<div class="ii_foot">
				<a href="#">相关链接</a>      
				<a href="#">客户留言</a>
				<a href="#">站长统计</a>
				<a href="#">版权归属</a>
				<%=t("webcopyright")%>
			</div>
		</div>
	</div>
</div>

</body>
</html>