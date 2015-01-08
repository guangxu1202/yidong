<%
' ------------------------------------------
' ESCMS V2.0 系统函数文件
' 请勿修改此文件任何信息,否则由此带来的麻烦请自己解决吧.
' 详细配置说明请查阅官方网站,请勿删下以下作者信息
' Web:http://www.escms.cn Email:escms@escms.cn
' ------------------------------------------

' ------------------------------------------
' 定义全局变量
' ------------------------------------------
Dim eConn,eRs,eConnStr,TempStr,sortStr,dRs
' 版本号 > 请勿修改
Const Versions = "JAMCMS V1.2 "


BrandNewDay 0
' ============================================
' 执行每天只需处理一次的事件 Compulsory＝1强制
' ============================================
Sub BrandNewDay(Compulsory)
	Dim sDate, y, m, d, w
	Dim sDateChinese
	sDate = Date()
	
	If Application("date_today") = sDate and Compulsory <> 1 Then Exit Sub

	y = CStr(Year(sDate))
	m = CStr(Month(sDate))
	If Len(m) = 1 Then m = "0" & m
	d = CStr(Day(sDate))
	If Len(d) = 1 Then d = "0" & d
	w = WeekdayName(Weekday(sDate))
	sDateChinese = y & "年" & m & "月" & d & "日&nbsp;" & w

	Application.Lock
	dbopen
	Application("date_today") = sDate
	Application("date_chinese") = sDateChinese		'今天的中文样式
	Application.Unlock
	eRs.open "select top 1 * from web_config",eConn,0,1
		Application("web_name") = eRs("webname")
		Application("web_host") = eRs("webhost")
		Application("web_keywords") = eRs("webkeywords")
		Application("web_description") = eRs("webdescription")
		Application("web_beian") = eRs("webbeian")
		Application("web_ishtml") = eRs("ishtml")
		Application("web_user") = eRs("webuser")
		Application("web_copyright") = eRs("webcopyright")
		Application("web_IsMsgAudit") = eRs("IsMsgAudit")
	eRs.Close
	dbclose
End Sub
' ------------------------------------------------------------
' 数据库相关函数,如果提示数据库连接失败,请修改config.asp!!此处不用修改
' -------------------------------------------------------------
Function DBopen()
	On Error Resume Next
	Set eConn = Server.CreateObject("ADODB.Connection")
	Set eRs = Server.CreateObject( "ADODB.Recordset" )
	Set dRs = Server.CreateObject( "ADODB.Recordset" )
	'err.Clear
	If DBType = "ACCESS" Then
		if Sysfolder <> "" then
			eConnStr = Server.MapPath("/" & Sysfolder & "/" & DBfolder & "/" & DBname)
		else
			eConnStr = Server.MapPath("/" & DBfolder & "/" & DBname)
		end if
		eConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & eConnStr & ";User Id=admin;Password=;"
		'eConnStr = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & eConnStr & ";Uid=Admin;Pwd=;"  '某些系统下不支持上面的连接方式,请使用这句
	ElseIf DBType = "SQL" Then
		eConnStr = SQLconnStr
	Else
		Set eConn = nothing
		Set eRs = nothing
		Set dRs = nothing
		Response.Write("数据库类型错误,请正确配置include/config.asp文件!<a href=""http://www.jamvs.com"" target=""_blank"">请联系网络管理员</a>")
		Response.End()
	End If
	eConn.open eConnStr
	If Err.Number <> 0 then
		eConn.Close
		Set eConn = nothing
		Set eRs = nothing
		Set dRs = nothing
		Response.Write("数据库连接错误!请正确配置include/config.asp文件!<a href=""http://www.jamvs.com"" target=""_blank"">请联系网络管理员</a>")
		Response.End()
	End If
End Function

' ----------------------------------------------------------
' 关闭数据库
' ----------------------------------------------------------
Function DBclose()
	On Error Resume Next
	eRs.Close
	eConn.Close
	Set eRs = nothing
	Set dRs = nothing
End Function

' ----------------------------------------------------------
' 获取用户来源IP
' ----------------------------------------------------------
Function UserIP() 
	dim Tempip
  Tempip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
  If Tempip = "" Then
    Tempip= Request.ServerVariables("REMOTE_ADDR") 
  End If
  UserIP = Tempip 
End Function 
' ----------------------------------------------------------
' 过滤HTML代码
' ----------------------------------------------------------
Function RemoveHTML(strHTML)
	on error resume next
	if inull(strHTML) = "" then
		RemoveHTML = ""
		exit function
	end if
	Dim objRegExp, Match, Matches
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = False
	objRegExp.Global = True
	objRegExp.Pattern = "(\<.[^\>]+\>)"
	Set Matches = objRegExp.Execute(strHTML)
	For Each Match in Matches
		strHtml=Replace(strHTML,Match.Value,"")
	Next
	RemoveHTML=strHTML
	Set objRegExp = Nothing
	if err.number <> 0 then RemoveHTML = ""
End Function
' ----------------------------------------------------------
' 获取字符串长度
' ----------------------------------------------------------
function getStringLen(str)
	on error resume next       
	dim l,c,i,t
	l=len(str)
	t=l
	for i=1 to l
		c=asc(mid(str,i,1))
		if c>=128 or c<0 then t=t+1
	next		
	getStringLen=t	
	if err.number<>0 then err.clear
end function

function getlenstr(str,length)
	if length = 0 then 
		getlenstr = str
		exit function
	else
		getlenstr = getSubString(RemoveHTML(str),Length)
	end if
end function
' ----------------------------------------------------------
' 截取字符串
' ----------------------------------------------------------
function getSubString(str,Length)
	on error resume next   
	if Length = 0 then 
		getSubString = ""
		exit function
	end if    
	dim l,c,i,hz,en
	l=len(str)
	if l<length then
		getSubString=str
	else
		hz=0
		en=0
		for i=1 to l
			c=asc(mid(str,i,1))
			if c>=128 or c<0 then 
				hz=hz+1
			else
				en=en+1
			end if
	
			if en\2+hz>=length then
				exit for
			end if
		next		
		getSubString=left(str,i) & "..."
	end if
	if err.number<>0 then err.clear
end function

' ----------------------------------------
' 获取当前页带参数的地址，可任意调用。
' 参数 parameter 页码参数,地址不返回此参数.以防重复
' ----------------------------------------
Function Url_address(pagestr)
	Dim uAddress,ItemUrl,Mitem,Get_Url,get_c,getstr,urladdress
	uAddress = CStr(Request.ServerVariables("SCRIPT_NAME"))
	ItemUrl = ""
	If (Request.QueryString <> "") Then
		uAddress = uAddress & "?"
		For Each Mitem In Request.QueryString
			If LCase(MItem) <> LCase(pagestr) then
				ItemUrl = ItemUrl & MItem &"="& trim(Server.URLEncode(Request.QueryString(""&MItem&""))) & "&"
			end if
		Next
	end if
	Get_Url = uAddress & ItemUrl
	if LCase(right(Get_Url,3))<>"asp" then
		Get_Url=left(Get_Url,len(Get_Url)-1)
	end if
	urladdress = Get_Url
	if Instr(urladdress,"?")>0 then
		urladdress=urladdress&"&"
	Else
		urladdress=urladdress&"?"
	end if
	url_address = urladdress
end function

' ---------------------------------------------------------------
' 分页代码函数
' 参数: pagecount:总页数，pageno:当前页
' 比如: response.write pageshow(ors.pagecount,request("page"))
' ---------------------------------------------------------------
Function WritePage(pagecounts,pageno)
	Dim TemplatePage,page2,s
	if pageno < 1 then pageno = 1
	if pageno > pagecounts then pageno = pagecounts
	TemplatePage = "<div id=""pagelist"">"  & vbcrlf
		if pageno=1 then
			TemplatePage = TemplatePage &  " <a href=""javascript:void(0)"" class=""unselected"">[首 页]</a> <a href=""javascript:void(0)"" class=""unselected"">[前一页]</a> " & vbcrlf
		else
			TemplatePage = TemplatePage &  " <a href=" & url_address("page") & "page=1"&">[首 页]</a> " & vbcrlf & " <a href=" & url_address("page") & "page="  & pageno-1 & ">[前一页]</a> " & vbcrlf
		end if
		page2=(pageno-(pageno mod 5))/5
		if page2<1 then page2=0
		for s=page2*5-1 to page2*5+5
			if s>0 then
				if s=cint(pageno) then 
					TemplatePage = TemplatePage &  " <a href=""#"" class=""selected"">["& s & "]</a>" & vbcrlf
				else
					if s=1 then
						TemplatePage = TemplatePage &  " <a href=" & url_address("page") & "page=1>["& s &"]</a>" & vbcrlf
					else
						TemplatePage = TemplatePage &  " <a href=" & url_address("page") & "page="&s&">["& s &"]</a>" & vbcrlf
					end if
				end if
				if s = Pagecounts then
					exit for
				end if
			end if
		next
		if cint(pageno) = Pagecounts then
			TemplatePage = TemplatePage &  " <a href=""javascript:void(0)""  class=""unselected"">[后一页]</a> <a href=""javascript:void(0)""  class=""unselected"">[尾 页]</a>" & vbcrlf
		else
			TemplatePage = TemplatePage &  " <a href="& url_address("page") & "page=" & pageno+1 &">[后一页]</a> " & vbcrlf & " <a href=" & url_address("page")  & "page=" & Pagecounts&">[尾 页]</a>" & vbcrlf 
		end if 
		TemplatePage = TemplatePage &"</div>"&vbcrlf
	WritePage = TemplatePage
End function

' ----------------------------------------------------------
' 一个加密的函数
' ----------------------------------------------------------
Function EnPas(CodeStr)
	Dim CodeLen,CodeSpace,NewCode,cecr,cecb,cec
	CodeLen = 30
	CodeSpace = CodeLen - Len(CodeStr)
	If Not CodeSpace < 1 Then
		For cecr = 1 To CodeSpace
			CodeStr = CodeStr & Chr(21)
		Next
	End If
	NewCode = 1
	Dim Been
	For cecb = 1 To CodeLen
		Been = CodeLen + Asc(Mid(CodeStr,cecb,1)) * cecb
		NewCode = NewCode * Been
	Next
	CodeStr = NewCode
	NewCode = Empty
	For cec = 1 To Len(CodeStr)
		NewCode = NewCode & CfsCode(Mid(CodeStr,cec,3))
	Next
	For cec = 20 To Len(NewCode) - 18 Step 2
		EnPas = EnPas & Mid(NewCode,cec,1)
	Next
		EnPas = "ES-" & EnPas
End Function
Function CfsCode(Word)
	Dim cc
	For cc = 1 To Len(Word)
		CfsCode = CfsCode & Asc(Mid(Word,cc,1))
	Next
	CfsCode = Hex(CfsCode)
End Function


' ----------------------------------------------------------
' 获取最新版本号,警告,修改此函数会影响您的正常使用
' ----------------------------------------------------------
Function GetNewVersion()
	GetNewVersion = "<script language=""javascript"" type=""text/javascript"" src=""http://www.escms.cn/update/newversion_new.asp?version=" & server.urlencode(Versions) & """></script>"
End Function

' ----------------------------------------------------------
' 管理后台显示系统名称  此名称可以修改为自己的名称
' ----------------------------------------------------------
function manageName()
	manageName = "<a href=""default.asp""><span>企业网站内容管理系统</span></a>"
end function

' ----------------------------------------------------------
' 管理后台页脚版权信息  免费版请保留此信息
' ----------------------------------------------------------
function manageCopyright()
	manageCopyright = "<p id=""footer""><a href=""http://www.jamvs.com"">技术支持</a></p>"
end function

' ----------------------------------------------------------
' 弹出一个提示信息并终止输出
' ----------------------------------------------------------
sub writealert(str,isreload)
	on error resume next
	oRs.Close : Set oRs = nothing
	conn.Close : set conn = nothing
	response.Write("<script language=""javascript"" type=""text/javascript"">" & vbcrlf)
	response.Write("alert('"&str&"');" & vbcrlf)
	if isreload = "1" then 
		response.Write("window.parent.location.reload();" & vbcrlf)
	elseif isreload <> "0" then
		response.Write("window.parent.location='"&isreload&"';" & vbcrlf)
	end if
	response.Write("</script>" & vbcrlf)
	response.End()
end sub

' ----------------------------------------------------------
' 一组数组,用于下拉框,权限名称等的输出...写成数组较方便
' ----------------------------------------------------------
function sorttypelist(tid)
	Dim listArray,arr,templist
	listArray = array("单页简介","新闻文章","图片产品","外部栏目","留言客服")
	for arr = 1 to ubound(listArray)+1
		if arr = tid then
			templist = templist & "<option value="""&arr&""" selected>"&listArray(arr-1)&"</option>"
		else
			templist = templist & "<option value="""&arr&""">"&listArray(arr-1)&"</option>"
		end if
	next
	sorttypelist = templist
end function
function sorttypename(tid)
	Dim listArray,arr,templist
	listArray = array("单页简介","新闻文章","图片产品","外部栏目","留言客服","下载中心")
	sorttypename = listArray(tid-1)
end function
function tagTypeList(tid)
	Dim listArray,arr,templist
	listArray = array("全局信息","栏目列表(导航或子栏目)","列表页通用","新闻列表","产品列表","留言列表","单页内容","新闻内容","产品内容","留言框样式","纯HTML")
	for arr = 1 to ubound(listArray)+1
		if arr = tid then
			templist = templist & "<option value="""&arr&""" selected>"&listArray(arr-1)&"</option>"
		else
			templist = templist & "<option value="""&arr&""">"&listArray(arr-1)&"</option>"
		end if
	next
	tagTypeList = templist
end function
function tagTypeName(tid)
	Dim listArray,arr,templist
	listArray = array("全局信息","栏目列表(导航或子栏目)","列表页通用","新闻列表","产品列表","留言列表","单页内容","新闻内容","产品内容","留言框样式","纯HTML")
	tagTypeName = listArray(tid-1)
end function
function userGroupList(tid)
	Dim listArray,arr,templist
	listArray = array("系统管理员","内容管理员","VIP用户","普通用户")
	for arr = 1 to ubound(listArray)+1
		if arr = tid then
			templist = templist & "<option value="""&arr*5&""" selected>"&listArray(arr-1)&"</option>"
		else
			templist = templist & "<option value="""&arr*5&""">"&listArray(arr-1)&"</option>"
		end if
	next
	userGroupList = templist
end function
function userGroupName(tid)
	Dim listArray,arr,templist
	listArray = array("系统管理员","内容管理员","VIP用户","普通用户")
	userGroupName = listArray(int(tid/5))
end function

' ----------------------------------------------------------
' 正则查找,返回第一个匹配项的第N个结果
' ----------------------------------------------------------
Function RegExpRet(strng,patrn,n)
  on error resume next
  Dim regEx, retVal ,RetStr,Matches ,Match
  Set regEx = New RegExp
  regEx.Pattern = patrn
  regEx.IgnoreCase = False
  set Matches  = regEx.Execute(strng)
  set Match = Matches(0)
  if err.number = 0 then
  	RetStr = Match.SubMatches(n)
  else
  	err.clear
  	RetStr = ""
  end if
  RegExpRet = RetStr
End Function
' ----------------------------------------------------------
' 正则匹配,匹配返回true,否则false
' ----------------------------------------------------------
Function RegExpTest(strng, patrn)
  Dim regEx, retVal
  Set regEx = New RegExp 
  regEx.Pattern = patrn
  regEx.IgnoreCase = False '不区分大小写
  retVal = regEx.Test(strng)
  If retVal Then
    RegExpTest = true
  Else
    RegExpTest = false
  End If
End Function

' ----------------------------------------------------------
' 正则替换
' ----------------------------------------------------------
Function ReplaceTest(str1,patrn, replStr)
  Dim regEx               ' 建立变量。
  Set regEx = New RegExp               ' 建立正则表达式。
  regEx.Global = True
  regEx.Pattern = patrn               ' 设置模式。
  regEx.IgnoreCase = True               ' 设置是否区分大小写。
  ReplaceTest = regEx.Replace(str1, replStr)         ' 作替换。
End Function

' ----------------------------------------------------------
' 解析自定义标签,愁死我了
' ----------------------------------------------------------
Function RegExpDiyTag(TempStr,diytagValue)
  Dim regEx, Match, Matches,RetStr
  Set regEx = New RegExp
  regEx.Pattern = "{diy:([^}]+)}"
  regEx.IgnoreCase = True 
  regEx.Global = True
  Set Matches = regEx.Execute(TempStr)
  RetStr = TempStr
  For Each Match in Matches
	RetStr = replace(RetStr,Match.Value,RegExpRet("$$$"&diytagValue, "\$\$\$" & trim(Match.SubMatches(0)) & "\|\|([^\$]+)\$\$\$",0))
  Next
  RegExpDiyTag = RetStr
End Function


' ----------------------------------------------------------
' 对入库的{进行解析
' ----------------------------------------------------------
function repreg(str)
	repreg = replace(replace(str,"{","@@@(@@@"),"}","@@@)@@@")
end function
' ----------------------------------------------------------
' 对}进行反解析
' ----------------------------------------------------------
function unrepreg(str)
	unrepreg = replace(replace(inull(str),"@@@(@@@","{"),"@@@)@@@","}")
end function

' ----------------------------------------------------------
' 把字符串进行HTML解码,替换server.htmlencode
' 去除Html格式，用于显示输出
' ----------------------------------------------------------
Function outHTML(str)
	Dim sTemp
	sTemp = str
	outHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
	sTemp = Replace(sTemp, Chr(10), "<br />")
	outHTML = sTemp
End Function


' ----------------------------------------------------------
' 获取当前页文件名 不含任何路径哦
' ----------------------------------------------------------
Function getFileName()
	Dim url
	url = Request.ServerVariables("SCRIPT_NAME")
	if instr(url,"/") > 0 then 
		url = mid(url,InStrRev(url,"/")+1)
	end if
	if instr(url,"?") > 0 then 
		url = mid(url,1,InStrRev(url,"?"))
	end if
	getFileName = url
End Function

' ----------------------------------------------------------
' 用户验证
' ----------------------------------------------------------
function checklogin(userpower)
	if request.Cookies(CookiesKey)("username") = "" then
		response.Write("<script language=""javascript"" type=""text/javascript"">" & vbcrlf)
		response.Write("alert('你还未登陆,请重新登陆!');window.location='login.asp?act=out';")
		response.Write("</script>" & vbcrlf)
		response.End()
	end if
	if int(request.Cookies(CookiesKey)("usergroup")) < 1 or int(userpower) < int(request.Cookies(CookiesKey)("usergroup")) then
		response.Write("<script language=""javascript"" type=""text/javascript"">" & vbcrlf)
		response.Write("alert('对不起,你没有操作此功能的权限,系统将自动返回上一页!');history.go(-1);")
		response.Write("</script>" & vbcrlf)
		response.End()
	end if
	dbopen
		eRs.open "select top 1 userid,username,userpass,usergroup from users where username = '"&replace(request.Cookies(CookiesKey)("username"),"'","")&"'",eConn,0,1
		if EnPas(eRs("username") & eRs("usergroup") & eRs("userpass")) <> request.Cookies(CookiesKey)("usercode") then
			dbclose
			response.Write("<script language=""javascript"" type=""text/javascript"">" & vbcrlf)
			response.Write("alert('登陆信息异常,请重新登陆!');window.location='login.asp?act=out';")
			response.Write("</script>" & vbcrlf)
			response.End()
		end if
		eRs.close
	dbclose
end function

' ----------------------
' 获取标签样式
' ----------------------
function getTagStr(tagName)
	on error resume next
	'response.Write("##########" & tagName & "####$")
	eRs.open "select top 1 tagcontent from templateTags where tagName = '"&tagName&"'",eConn,0,1
	if not eRs.eof then 
		getTagStr = inull(eRs(0))
	else
		getTagStr = ""
	end if
	'response.Write(err.description)
	eRs.Close
	if err.number <> 0 then getTagStr = ""	
end function

' ============================================
' 格式化时间(显示)
' 参数：n_Flag
'	1:"yyyy-mm-dd hh:mm:ss"
'	2:"yyyy-mm-dd"
'	3:"hh:mm:ss"
'	4:"yyyy年mm月dd日"
'	5:"yyyymmdd"
'   6:"mm-dd"
' ============================================
Function Format_Time(s_Time, n_Flag)
	Dim y, m, d, h, mi, s
	Format_Time = ""
	If IsDate(s_Time) = False Then Exit Function
	y = cstr(right("00" & year(s_Time),2))
	m = cstr(right("00" & month(s_Time),2))
	d = cstr(right("00" & day(s_Time),2))
	h = cstr(right("00" & hour(s_Time),2))
	mi = cstr(right("00" & minute(s_Time),2))
	s = cstr(right("00" & second(s_Time),2))

	Select Case n_Flag
	Case 1
		' yyyy-mm-dd hh:mm:ss
		Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
		' yyyy-mm-dd
		Format_Time = y & "-" & m & "-" & d
	Case 3
		' hh:mm:ss
		Format_Time = h & ":" & mi & ":" & s
	Case 4
		' yyyy年mm月dd日
		Format_Time = y & "年" & m & "月" & d & "日"
	Case 5
		' yyyymmdd
		Format_Time = y & m & d
	Case 6
		Format_Time = m & "-" & d
	case else
		Format_Time = y & "-" & m & "-" & d
	End Select
End Function
' ----------------------------------------------------------
' FormatTime 改进的Time格式化,EF
' s_Time 要格式化的日期
' s_Flag 格式,可以用如 "yyyy-mm-dd" "mm-dd"类似的自定义格式 [yyyy mm dd hh nn ss 年月日时分秒]
' 示例 FormatTime(date(),"yy-mm-dd hh:nn:ss")
' ----------------------------------------------------------
function FormatTime(s_Time,s_Flag)
	Dim y, m, d, h, mi, s, ReturnStr
	FormatTime = ""
	If IsDate(s_Time) = False Then Exit Function
	y = year(s_Time)
	m = month(s_Time)
	d = day(s_Time)
	h = hour(s_Time)
	mi = minute(s_Time)
	s = second(s_Time)
	
	ReturnStr = lcase(s_Flag)
	ReturnStr = replace(ReturnStr,"yyyy",y)
	ReturnStr = replace(ReturnStr,"yy",right(y,2))
	ReturnStr = replace(ReturnStr,"mm",right("00" & m,2))
	ReturnStr = replace(ReturnStr,"m",m)
	ReturnStr = replace(ReturnStr,"dd",right("00" & d,2))
	ReturnStr = replace(ReturnStr,"d",d)
	ReturnStr = replace(ReturnStr,"hh",right("00" & h,2))
	ReturnStr = replace(ReturnStr,"h",h)
	ReturnStr = replace(ReturnStr,"nn",right("00" & mi,2))
	ReturnStr = replace(ReturnStr,"n",mi)
	ReturnStr = replace(ReturnStr,"ss",right("00" & s,2))
	ReturnStr = replace(ReturnStr,"s",s)
	FormatTime = ReturnStr
end function

' ----------------------------------------------------------
' 如果值为null返回空,以处理null值带来的错误
' ----------------------------------------------------------
function inull(s)
	dim str
	if isnull(s) then
	 str = ""
	else
	 str = s
	end if
	inull = str
end function
' ----------------------------------------------------------
' 如果值为null返回0,以处理null值带来的错误
' ----------------------------------------------------------
function snull(s)
	dim str
	if isnull(s) then
	 str = 0
	else
	 str = s
	end if
	snull = str
end function
' ----------------------------------------------------------
' 强暴的转为数字,如果不能转,就返回0,以处理错误类型引起的错误
' ----------------------------------------------------------
function eint(s)
	dim str
	if isnull(s) or s = "" then
	 str = 0
	else
	 on error resume next
	 str = int(s)
	 if err.number <> 0 then
	 	err.clear
		str = 0
	 end if
	end if
	eint = str
end function
' ----------------------------------------------------------
' 
' ----------------------------------------------------------
%>
