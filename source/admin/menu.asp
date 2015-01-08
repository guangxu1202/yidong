<%
function iscadmin()
	if int(request.Cookies(CookiesKey)("usergroup")) > 1 and int(request.Cookies(CookiesKey)("usergroup")) < 6 then 
		iscadmin = true
	else
		iscadmin = false
	end if
end function 
function manageMenu(mNums,Permissions)  '当前项,权限
	Dim tempStr
	tempStr = tempStr & "<li><a href=""default.asp"" {0}>管理首页</a></li>"&vbcrlf
	tempStr = tempStr & "<li><a href=""alone.asp"" {1}>网站内容管理</a></li>"&vbcrlf
    if iscadmin then tempStr = tempStr & "<li><a href=""sort.asp"" {2}>栏目管理</a></li>"&vbcrlf	
    if iscadmin then tempStr = tempStr & "<li><a href=""system.asp"" {3}>系统管理</a></li>"&vbcrlf
    'if request.Cookies(CookiesKey)("usergroup")) > 1 and int(request.Cookies(CookiesKey)("usergroup") < 6 then tempStr = tempStr & "<li><a href=""#"" {4}>插件管理</a></li>"&vbcrlf
    tempStr = tempStr & "<li class=""logout""><a href=""login.asp?act=out"">退出</a></li>"&vbcrlf
	tempStr = tempStr & "<li class=""logout""><a href=""../"" target=""_blank"">网站首页</a></li>"&vbcrlf
	tempStr = replace(tempStr,"{"&mNums&"}","class=""active""")
	tempStr = ReplaceTest(tempStr,"\{\d{1,2}\}","")
	manageMenu = tempStr	
end function
function smallMenu(tid,mNums)
	Dim tempStr
	select case tid
	case 1
		if iscadmin then tempStr = tempStr & "<li><a href=""system.asp"" {1}>系统配置</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""message.asp"" {2}>留言管理</a></li>" & vbcrlf
		'if iscadmin then tempStr = tempStr & "<li><a href=""webcounts.asp"" {3}>查看流量</a></li>" & vbcrlf
		if iscadmin then tempStr = tempStr & "<li><a href=""tagstyle.asp"" {4}>标签样式管理</a></li>" & vbcrlf
		if iscadmin then tempStr = tempStr & "<li><a href=""users.asp"" {5}>用户管理</a></li>" & vbcrlf
		'if iscadmin then tempStr = tempStr & "<li><a href=""database.asp"" {6}>数据库管理</a></li>" & vbcrlf
		'if iscadmin then tempStr = tempStr & "<li><a href=""templates.asp"" {7}>模板管理</a></li>" & vbcrlf
		'if iscadmin then tempStr = tempStr & "<li><a href=""fileups.asp"" {8}>上传文件管理</a></li>" & vbcrlf
	case 2
		if iscadmin then tempStr = tempStr & "<li><a href=""sort.asp#add"" {1}>添加栏目</a></li>" & vbcrlf
		if iscadmin then tempStr = tempStr & "<li><a href=""sort.asp"" {2}>栏目管理</a></li>" & vbcrlf
	case 3
		tempStr = tempStr & "<li><a href=""alone.asp"" {1}>单页管理</a></li>" & vbcrlf
		tempStr = tempStr & "<li><a href=""news.asp"" {2}>新闻(文章)管理</a></li>" & vbcrlf
		tempStr = tempStr & "<li><a href=""product.asp"" {3}>产品管理</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""mainProduct.asp"" {5}>主打产品管理</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""range.asp"" {6}>练习场地管理</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""hotel.asp"" {7}>酒店信息管理</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""tips.asp"" {8}>提示框修改</a></li>" & vbcrlf
		tempStr = tempStr & "<li><a href=""message.asp"" {4}>留言管理</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""job.asp"" {5}>人才管理</a></li>" & vbcrlf
	case else
		tempStr = tempStr & "<li><a href=""default.asp"" {1}>WelCome</a></li>" & vbcrlf
		tempStr = tempStr & "<li><a href=""news.asp#add"" {2}>添加新闻</a></li>" & vbcrlf
		tempStr = tempStr & "<li><a href=""alone.asp"" {3}>管理单页</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""product.asp#add"" {4}>添加产品</a></li>" & vbcrlf
		tempStr = tempStr & "<li><a href=""message.asp"" {5}>查看留言</a></li>" & vbcrlf
		'tempStr = tempStr & "<li><a href=""http://www.efsys.cn/help.html"" {5}>使用帮助</a></li>" & vbcrlf
	end select
	tempStr = replace(tempStr,"{"&mNums&"}","class=""active""")
	tempStr = ReplaceTest(tempStr,"\{\d{1,2}\}","")
	smallMenu = tempStr

end function

			
%>