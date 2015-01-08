<%
dim sortids

sortids=t("sortid")
'response.Write(sortids)
'response.End()
%>


<div class="header">
	<div class="center">
		<div class="logo"><a href="default.asp"></a></div>
        <div class="con"><a href="#" target="_blank">联系方式</a> | <a href="#">中/英</a></div>
		<div class="headR">
        	<%=T("栏目$一级导航栏目")%>
			<div class="search">
				<form action="newssearch.asp" id="form1" name="form1" method="get">
					<input class="inputText" type="text" name="keyword" value=""/>
					<input class="inputSubmit" type="submit" value=""/>
				</form>
			</div>
		</div>
	</div>
</div>

<%
'pic change
dim banner,ico
select case getsortID()
	
case 42
	banner = "banner07"
	ico = "ico0005.png"
case 43
	banner = "banner06"
	ico = "ico0005.png"
case 44
	banner = "banner08"
	ico = "ico0005.png"
case 45
	banner = "banner09"
	ico = "ico0005.png"
case 46
	banner = "banner04"
	ico = "ico0005.png"
case 51,52,53,54,32						
	banner = "banner05"
	ico = "ico0004.png"
case 33,55,56,58					
	banner = "banner10"
	ico = "ico0006.png"
case 29,38,39,40,41					
	banner = "banner02"
	ico = "ico0003.png"
case 31,47,48,49,50					
	banner = "banner03"
	ico = "ico0002.png"
case 28,34,35,36,37					
	banner = "banner01"
  ico = "ico0001.png"
case else
  banner = "banner02"
	ico = "ico0003.png"
end select
%>