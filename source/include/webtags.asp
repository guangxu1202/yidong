<%
' ===================================
' ESCMS V2.0 标签
' 此文件可根据需求修改
' Author : EF Mail: escms@escms.cn
' ===================================


' ----------------------
' 获取当前栏目ID
' ----------------------
function getSortID()
	dim sortid
	sortid = request.QueryString("sortid")
	if sortid = "" then
	  on error resume next
	  sortid = eConn.execute("select top 1 sortid from sorts where sortlisttemplates = '"&getFileName()&"'" )(0)
	  if err.number <> 0 then
			err.clear
		  sortid = 0
	  end if
	end if
	sortid = int(snull(sortid))
	getSortID = sortid
end function

' ----------------------
' 获取栏目的顶层ID
' ----------------------
function getsortTID(sid)
	dim gRs
	set gRs = eConn.execute("select top 1 sortid,sortfid,sortlisttemplates from sorts where sortid = " & sid)
	if gRs.eof then
		getsortTID = 0
		exit function
	else
		if int(gRs("sortfid")) = 0 then
			getsortTID = int(snull(gRs("sortid")))
			exit function
		else
			getsortTID = getsortTID(snull(gRs("sortfid")))
		end if
	end if
	
end function

' ----------------------
' 获取栏目的上层ID
' ----------------------
function getsortFID(sid)
	dim gRs
	set gRs = eConn.execute("select top 1 sortid,sortfid,sortlisttemplates from sorts where sortid = " & sid)
	if gRs.eof then
		getsortFID = 0
		exit function
	else
		getsortFID = snull(gRs("sortid"))
		exit function
	end if 
end function

' ----------------------
' 获取当前栏目名称或ID对应名称
' sid为栏目ID,不过ID未指定时,可以根据sortid来获取,或者文件名和栏目中的相同,这都没有,就空着回去吧
' ----------------------
function getSortName(sid)
	dim sortid
	if sid = 0 then 
		sortid = getSortID
	else
		sortid = sid
	end if
	if sortid = 0 then
		getSortName = ""
	else
		on error resume next
		getSortName = eConn.execute("select sortname from sorts where sortid = " & int(sortid))(0)
		if err.number <> 0 then
			err.clear
			getSortName = ""
		end if
	end if
end function

' ----------------------
' 获取单页内容  '此函数建议用T代替
' aid: 要获取的单页ID
' Info: 摘要 or 内容 or 图片
' sizes: 字数 0 为全部
' ishtml: 1 为显示html,0 为纯文本
' ----------------------
function GetAlone(aid,Info,sizes,ishtml)
	Dim tempTagStr,returnStr,tempValue,tempPage,tempcontent,tempAbstract,aloneid
	'---判断要获取的ID,参数aid优先,然后是提交的aloneid,再是sortid,要是都没有,那就0吧,取第一条单页,总不能空着回去吧..
	if aid = 0 then  
		if request.QueryString("aloneid") = "" then  
			if request.QueryString("sortid") = "" then
				aloneid = 0
			else
				on error resume next
				aloneid = eConn.execute( "select top 1 aloneid from alones where aloneissort = " &getSortID )(0)
				if err.number <> 0 then 
					err.clear
					aloneid = 0
				else
					aloneid = int(aloneid)
				end if
			end if
		else
			aloneid = int(request.QueryString("aloneid"))
		end if
	else
		aloneid = aid
	end if
	'--判断结束-------------
	select case Info
	case "摘要"
		tempTagStr = getTagStr("单页摘要")
		if tempTagStr = "" then tempTagStr = "{单页摘要}"
	case "内容"
		tempTagStr = getTagStr("单页内容")
		if tempTagStr = "" then tempTagStr = "{单页内容}"
	case "图片"
		tempTagStr = getTagStr("单页图片")
		if tempTagStr = "" then tempTagStr = "{单页图片}"
	end select
	if not isnumeric(aloneid) then
		eRs.open "select top 1 * from alones where alonename = '" & aloneid & "' ",eConn,0,1
	else
		eRs.open "select top 1 * from alones where aloneid >= " & aloneid & " order by aloneid asc",eConn,0,1
	end if
	if not eRs.eof then 
		tempPage =  eRs("aloneshowpage") 
		if tempPage = "" then tempPage = "about.asp"
		returnStr = replace(tempTagStr,"{单页标题}",eRs("alonename"))
		returnStr = replace(returnStr,"{单页ID}",eRs("aloneid"))
		returnStr = replace(returnStr,"{单页地址}",eRs("aloneshowpage") & "?sortid=" & eRs("aloneid"))
		returnStr = replace(returnStr,"{单页图片}",eRs("alonepic"))
		if ishtml = 0 then
			tempcontent = trim(outHTML(RemoveHTML(eRs("alonecontent"))))
			tempAbstract = trim(outHTML(RemoveHTML(eRs("aloneabstract")))) 'eRs("aloneabstract")
		else
			tempcontent = trim(eRs("alonecontent"))
			tempAbstract = trim(eRs("aloneabstract")) 'eRs("aloneabstract")
		end if
		if sizes <> 0 then
			tempcontent = getSubString(tempcontent,sizes)
			tempAbstract = getSubString(tempAbstract,sizes)
		end if
		returnStr = replace(returnStr,"{栏目名称}",T("栏目名称"))
		returnStr = replace(returnStr,"{栏目ID}",T("栏目ID"))
		returnStr = replace(returnStr,"{单页摘要}",tempAbstract)
		returnStr = replace(returnStr,"{单页内容}",tempcontent)
	else
		returnStr = tempTagStr
	end if
	eRs.Close
	GetAlone = unrepreg(returnStr)
end function

' ----------------------------
' 获取一级栏目
' ----------------------------
function GetNav(style)
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,sortid,isselect,tempListStr,tempsubStr
	'response.write T("sid") & "-"
	if T("sid") <> 0 then
		sortid = getSortTid(T("sid"))  '获取顶级栏目ID,用以控制一级导航当前栏目样式
	elseif tSortID = 0 then
		sortid = tSortID
	else
		sortid = 0
	end if
	isselect = false
	tempTagStr = getTagStr(style)
	if tempTagStr = "" then
		tempTagStr = "<ul class='nav'>[LIST]<li {selected}><a href=""{连接地址}"">{栏目名称}</a>"&GetSort("")&"</li>[/LIST]</ul>"
	end if	
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if
	
	returnStr = startStr
	eRs.open "select sortid,sortname,sortlisttemplates,sorttype from sorts where sortfid = 0 and sortisnav = true order by sortorder asc,sortid asc",eConn,0,1
	do while not eRs.eof
		tempsubStr = ""
		if cstr(eRs("sortid")) = cstr(sortid) then
			isselect = true
			tempListStr = replace(listStr,"{selected}"," class = ""selected"" ")
		else
			tempListStr = replace(listStr,"{selected}","")
		end if
		if eRs("sorttype") =4 then
			tempListStr = replace(tempListStr,"{连接地址}",eRs("sortlisttemplates"))
		else
			tempListStr = replace(tempListStr,"{连接地址}",eRs("sortlisttemplates") & "?sortid=" & eRs("sortid"))
		end if
		
		tempListStr = replace(tempListStr,"{栏目名称}",eRs("sortname"))
		tempListStr = replace(tempListStr,"{栏目ID}",eRs("sortid"))
			dRs.open "select sortid,sortname,sortlisttemplates from sorts where sortfid = "&eRs("sortid")&" and sortislist = true order by sortorder asc,sortid asc",eConn,1,1
			if not dRs.eof then
			tempsubStr = "<ol>"
			do while not dRs.eof
			
			
			tempsubStr = tempsubStr &"<li> <a href='"&dRs("sortlisttemplates")& "?sortid=" & dRs("sortid")&"' target='_self'>"& dRs("sortname")&"</a></li>"
			dRs.movenext
			loop
			tempsubStr = tempsubStr & "</ol>"
			end if
			dRs.close
		tempListStr = replace(tempListStr,"{一级导航子菜单}",tempsubStr)
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	if isselect = false then
		returnStr = replace(returnStr,"{selected}"," class = ""selected"" ")
	else
		returnStr = replace(returnStr,"{selected}","")
	end if
	returnStr = returnStr & endStr
	eRs.Close
	'替换全局
	returnStr = replaceGlobal(returnStr)
	GetNav = returnStr
end function

' ----------------------------
' 获取当前栏目的子栏目
' islocal 如果是0，当无子栏目时，调用同级目录 ，非0，就不调用同级子目录 
' ----------------------------
function GetSort(style,sid,islocal)
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,sortid,isselect,tempListStr
	if sid = 0 then
		sortid = getSortID
		if sortid = 0 then sortid = tSortID
	else
		sortid = sid
	end if
	isselect = false
	tempTagStr = getTagStr(style)
	if tempTagStr = "" then
		tempTagStr = "<ul class='subnav'>[LIST]<li {selected}><a href=""{连接地址}"">{栏目名称} &gt;</a></li>[/LIST]</ul>"
	end if	
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if
	returnStr = startStr
	eRs.open "select sortid,sortname,sortlisttemplates from sorts where sortfid = "&sortid&" and sortislist = true order by sortorder asc,sortid asc",eConn,0,1
	if eRs.eof then
		if islocal = 0 then
		eRs.Close
		eRs.open "select sortid,sortname,sortlisttemplates from sorts where sortfid = (select top 1 sortfid from sorts where sortid = "&sortid&") and sortislist = true order by sortorder asc,sortid asc",eConn,0,1
		end if
	end if
	do while not eRs.eof
		if cstr(eRs("sortid")) = cstr(sortid) then
			isselect = true
			tempListStr = replace(listStr,"{selected}"," class=""selected"" ")
		else
			tempListStr = replace(listStr,"{selected}","")
		end if
		tempListStr = replace(tempListStr,"{navlink}",eRs("sortlisttemplates") & "?sortid=" & eRs("sortid"))
		tempListStr = replace(tempListStr,"{连接地址}",eRs("sortlisttemplates") & "?sortid=" & eRs("sortid"))
		tempListStr = replace(tempListStr,"{navname}",eRs("sortname"))
		tempListStr = replace(tempListStr,"{栏目名称}",eRs("sortname"))
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	if isselect = false then
		returnStr = replace(returnStr,"{selected}"," class = ""selected"" ")
	else
		returnStr = replace(returnStr,"{selected}","")
	end if
	returnStr = returnStr & endStr
	eRs.Close
	returnStr = replaceGlobal(returnStr)
	GetSort = returnStr

end function

' ----------------------------
' ?? 这个？
' ----------------------------
function GetAloneList(style)
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,sortid,isselect,tempListStr
	sortid = getSortID
	isselect = false
	tempTagStr = getTagStr(style)
	if tempTagStr = "" then
		tempTagStr = "<ul>[LIST]<li {selected}><a href=""{navlink}"">{navname}</a></li>[/LIST]</ul>"
	end if	
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if
	returnStr = startStr
	eRs.open "select sortid,sortname,sortlisttemplates from sorts where sortfid = "&sortid&" and sortislist = true order by sortorder asc,sortid asc",eConn,0,1
	if eRs.eof then
		eRs.Close
		eRs.open "select sortid,sortname,sortlisttemplates from sorts where sortfid = (select top 1 sortfid from sorts where sortid = "&sortid&") and sortislist = true order by sortorder asc,sortid asc",eConn,0,1
	end if
	do while not eRs.eof
		if cstr(eRs("sortid")) = cstr(sortid) then
			isselect = true
			tempListStr = replace(listStr,"{selected}"," class=""selected"" ")
		else
			tempListStr = replace(listStr,"{selected}","")
		end if
		tempListStr = replace(tempListStr,"{navlink}",eRs("sortlisttemplates") & "?sortid=" & eRs("sortid"))
		tempListStr = replace(tempListStr,"{navname}",eRs("sortname"))
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	if isselect = false then
		returnStr = replace(returnStr,"{selected}"," class=""selected"" ")
	else
		returnStr = replace(returnStr,"{selected}","")
	end if
	returnStr = returnStr & endStr
	eRs.Close
	GetAloneList = returnStr

end function

' ----------------------
' 新闻文章列表函数
' sid 为栏目ID
' style 样式名称
' listnums 调用条数,如果是终级列表,则为每页数量
' dateformat 自定义日期格式
' titlesize 标题显示长度 中文一个算一个,英文俩个算一个
' issubclass 是否包含子类新闻 1  包含,0不包含
' Abstractsize 0 为不显示摘要,大于0为显示摘要的字数,这个始终是纯文本输出的
' ispiconly 是否仅显示图片文章,1为仅显示图片的,0为全部  '此选项有助于用户方便输出图片列表,图片大小,可在后台建个样式
' isviewpage 是否显示分页 1 显示 0 不显示
' ----------------------
function newsList(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
	dim intpage,sortid,pagesizes,pagecounts,pagesql,recordcounts,wheresql,fidsql,linenums,keyword
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,tempPage,tempListStr,listpage
	pagesizes = listnums
	'处理栏目ID,如果有输入,直接赋值,没有,则获取sortid,如果还没有,则根据文件名获取,还是没有?那就全部吧..
	if sid <> 0 then
		sortid = sid
	else
		sortid = getSortID
	end if
	'处理分页
	intpage = request.QueryString("page")
	if intpage = "" then
		intpage = 1
	else 
		intpage = int(intpage)
	end if
	'处理条件语句
	keyword = replace(request.QueryString("keyword"),"'","")
	if sortid <> 0 then
		wheresql = " and artsort = " & sortid
	end if
	if issubclass = "y" and sortid <> 0 then
		wheresql = " and artsort in (select sortid from sorts where sortid = "&sortid&" or sortid in (select sortid from sorts where sortfid="&sortid&" )  or sortfid in (select sortid from sorts where sortfid="&sortid&" )) "
	end if
	if ispiconly = "y" then
		wheresql = wheresql & " and artpic <> '' "
	end if
	keyword = replace(request.QueryString("keyword"),"'","")
	if keyword <> "" then
		wheresql = wheresql & " and arttitle like '%"&keyword&"%' "
	end if
	if right(ordertype,2) = "点击" then
		select case lcase(left(ordertype,1))
		case "年" '年
			wheresql = wheresql & " and datediff('yyyy',artupdatetime,'"&now()&"') = 0 "
		case "季" '季
			wheresql = wheresql & " and datediff('q',artupdatetime,'"&now()&"') = 0 "
		case "月" '月
			wheresql = wheresql & " and datediff('m',artupdatetime,'"&now()&"') = 0 "
		case "周" '周
			wheresql = wheresql & " and datediff('ww',artupdatetime,'"&now()&"') = 0 "
		case else
			wheresql = wheresql & " and datediff('m',artupdatetime,'"&now()&"') = 0 "
		end select 
	end if
	'处理分页
	if isviewpage = "y" then
		pagesql = "select count(*) from articles where  artisrec = false " & wheresql
		'response.Write(pagesql)
		recordcounts = eConn.execute(pagesql)(0)
		if recordcounts/pagesizes > int(recordcounts/pagesizes) then 
			pagecounts = int(recordcounts/pagesizes) + 1
		else
			pagecounts = int(recordcounts/pagesizes)
		end if
		if pagecounts < 1 then pagecounts = 1
	else
		intpage = 1
		pagecounts = 1	
	end if
	if ordertype = "置顶" then
		ordertype = " artistop desc,artupdatetime desc "
	elseif right(ordertype,2) = "点击" then
		ordertype = " artclicks desc,artupdatetime desc "
	elseif ordertype = "更新" then
		ordertype = " artupdatetime desc "
	elseif ordertype = "发布" then
		ordertype = " artaddtime desc "
	else
		ordertype = " artupdatetime desc "
	end if
	'处理SQL
	if intpage > 1 then
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (articles a left join sorts b on a.artsort = b.sortid) where artid not in(select top "&(intpage-1)*pagesizes &" artid from articles where artisrec = false "&wheresql&" order by "& ordertype &") and artisrec = false  "&wheresql&" order by "& ordertype &""
	else
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (articles a left join sorts b on a.artsort = b.sortid) where artisrec = false  "&wheresql&"  order by "& ordertype &""
	end if

	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "<ul>[LIST]<li><em>{更新时间}</em><a href=""{文章地址}"">{标题}</a></li>[/LIST]</ul>"
	end if
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if

	'response.Write("--------"&pagesql&"----------")
	startStr = replace(startStr,"{栏目名称}",getSortName(sortid))
	endStr = replace(endStr,"{栏目名称}",getSortName(sortid))

	startStr = replace(startStr,"{栏目ID}",sortid)
	endStr = replace(endStr,"{栏目ID}",sortid)
	
	returnStr = startStr
	'response.Write(pagesql)
	eRs.open pagesql,eConn,0,1
	do while not eRs.eof
		linenums = linenums + 1
		tempListStr = listStr
		tempPage = eRs("sortshowtemplates")
		listpage = eRs("sortlisttemplates")
		if tempPage = "" or isnull(tempPage) then tempPage = "newsshow.asp"
		if listpage = "" or isnull(listpage) then listpage = "newslist.asp"
		'{FH:2:twostyle}
		if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
			if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
			else
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
			end if
			
		end if
		tempListStr = replace(tempListStr,"{i}",linenums)
		tempListStr = replace(tempListStr,"{标题}",getlenstr(eRs("arttitle"),titlesize))
		tempListStr = replace(tempListStr,"{文章地址}",tempPage & "?newsid=" & eRs("artid"))
		tempListStr = replace(tempListStr,"{文章ID}",eRs("artid"))
		tempListStr = replace(tempListStr,"{来源}",eRs("artsource"))
		tempListStr = replace(tempListStr,"{作者}",eRs("artauthor"))
		tempListStr = replace(tempListStr,"{文章点击数}",eRs("artclicks"))
		tempListStr = replace(tempListStr,"{栏目名称}",inull(eRs("sortname")))
		tempListStr = replace(tempListStr,"{栏目地址}",listpage & "?sortid=" & eRs("artsort"))
		tempListStr = replace(tempListStr,"{图片地址}",eRs("artpic"))
		tempListStr = replace(tempListStr,"{摘要}",getSubString(eRs("artdescription"),Abstractsize))		
		tempListStr = replace(tempListStr,"{更新时间}",formatTime(eRs("artupdatetime"),dateformat))
		'tempListStr = replace(tempListStr,"{uptime}",Format_Time(eRs("artupdatetime"),dateformat))
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	eRs.Close
	returnStr = returnStr & endStr
	'替换全局标签
	returnStr = replaceGlobal(returnStr)
	returnStr = unrepreg(returnStr)

	if isviewpage = "y" then
		if instr(returnStr,"{分页}") > 0 then
			returnStr = replace(returnStr,"{分页}",writepage(pagecounts,intpage))
		else
			returnStr = returnStr & writepage(pagecounts,intpage)
		end if 
	end if
	newsList = returnStr
end function


' ----------------------
' 新闻内容调用
' ----------------------
function newsContent(style)
	Dim tempTagStr,returnStr,tempValue,intpage,breakpage,newsid,pagesplit
	breakpage = "<div style=""page-break-after: always""><span style=""display: none"">&nbsp;</span></div>"
	intpage = request.QueryString("page")
	newsid = request.QueryString("newsid")
	if not isnumeric(intpage) then 
		intpage = 1
	else
		intpage = int(intpage)
	end if
	if not isnumeric(newsid) then
		newsContent = "无效ID"
		exit function
	else
		newsid = int(newsid)
	end if
	if intpage < 1 then intpage = 1
	if newsid < 1 then exit function
	
	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "<h1>{标题}</h1><div id=""newscontent"">{文章内容}</div>"
	end if
	'response.Write("select top 1 * from articles from artisrec = false and artid = " & newsid)
	eRs.open "select top 1 * from articles where artisrec = false and artid = " & newsid,eConn,0,1
	if eRs.eof then
		newsContent = "无效ID"
		exit function
	end if
	
	pagesplit = split(ers("artcontent") & " ",breakpage)
	if intpage > ubound(pagesplit)+1 then intpage = ubound(pagesplit)+1
	
		tempListStr = replace(tempListStr,"{标题}",getlenstr(eRs("arttitle"),titlesize))
		tempListStr = replace(tempListStr,"{文章地址}",tempPage & "?newsid=" & eRs("artid"))
		tempListStr = replace(tempListStr,"{文章ID}",eRs("artid"))
		tempListStr = replace(tempListStr,"{来源}",eRs("artsource"))
		tempListStr = replace(tempListStr,"{作者}",eRs("artauthor"))
		tempListStr = replace(tempListStr,"{文章点击数}",eRs("artclicks"))
		'tempListStr = replace(tempListStr,"{栏目名称}",inull(eRs("sortname")))
		'tempListStr = replace(tempListStr,"{栏目地址}",listpage & "?sortid=" & eRs("artsort"))
		tempListStr = replace(tempListStr,"{栏目名称}",T("栏目名称"))
		tempListStr = replace(tempListStr,"{栏目ID}",T("栏目ID"))
		tempListStr = replace(tempListStr,"{图片地址}",eRs("artpic"))
		tempListStr = replace(tempListStr,"{摘要}",getSubString(eRs("artdescription"),Abstractsize))		
		tempListStr = replace(tempListStr,"{更新时间}",formatTime(eRs("artupdatetime"),dateformat))
		tempTagStr = replace(tempTagStr,"{文章内容}",pagesplit(intpage-1))
		tempListStr = RegExpDiyTag(tempListStr,eRs("prodiyvalue"))

	eRs.Close
	returnStr = replaceGlobal(returnStr)
	returnStr = unrepreg(tempTagStr)
	if ubound(pagesplit) > 0 then
		if instr(returnStr,"{分页}") > 0 then
			returnStr = replace(returnStr,"{分页}",writepage(ubound(pagesplit)+1,intpage))
		else
			returnStr = returnStr & writepage(ubound(pagesplit)+1,intpage)
		end if 
	end if

	newsContent = returnStr
	
end function

' ----------------------
' 产品列表函数  这个和新闻的差不多,是为了方便用户调用
' sid 为栏目ID
' style 样式名称
' listnums 调用条数,如果是终级列表,则为每页数量
' dateformat 日期格式,'	1:"yyyy-mm-dd hh:mm:ss" '	2:"yyyy-mm-dd" '	3:"hh:mm:ss" '	4:"yyyy年mm月dd日" '	5:"yyyymmdd" '   6:"mm-dd"
' titlesize 标题显示长度 中文一个算俩个,英文一个算一个
' issubclass 是否包含子类新闻 1  包含,0不包含
' Abstractsize 0 为不显示摘要,大于0为显示摘要的字数,这个始终是纯文本输出的
' ispiconly 是否仅显示有小图片文章,1为仅显示图片的,0为全部  '此选项有助于用户方便输出图片列表,图片大小,可在后台建个样式
' isviewpage 是否显示分页 1 显示 0 不显示
' ----------------------
function proList(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
	dim intpage,sortid,pagesizes,pagecounts,pagesql,recordcounts,wheresql,fidsql,linenums,keyword
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,tempPage,tempListStr,listpage
	pagesizes = listnums
	'处理栏目ID,如果有输入,直接赋值,没有,则获取sortid,如果还没有,则根据文件名获取,还是没有?那就全部吧..
	if sid <> 0 then
		sortid = sid
	else
		sortid = getSortID
	end if
	'处理分页
	intpage = request.QueryString("page")
	if intpage = "" then
		intpage = 1
	else 
		intpage = int(intpage)
	end if
	'处理条件语句
	if sortid <> 0 then
		wheresql = " and prosort = " & sortid
	end if
	if issubclass = "y" and sortid <> 0 then
		wheresql = " and prosort in (select sortid from sorts where sortid = "&sortid&" or sortid in (select sortid from sorts where sortfid="&sortid&" ) or sortfid in (select sortid from sorts where sortfid="&sortid&" )) "
	end if
	if ispiconly = "y" then
		wheresql = wheresql & " and prosmallpic <> '' "
	end if
	keyword = replace(request.QueryString("keyword"),"'","")
	if keyword <> "" then
		wheresql = wheresql & " and proname like '%"&keyword&"%' "
	end if
	if left(ordertype,2) = "点击" then
		select case lcase(right(ordertype,1))
		case "y" '年
			wheresql = wheresql & " and datediff('yyyy',proupdatetime,'"&now()&"') = 0 "
		case "j" '季
			wheresql = wheresql & " and datediff('q',proupdatetime,'"&now()&"') = 0 "
		case "m" '月
			wheresql = wheresql & " and datediff('m',proupdatetime,'"&now()&"') = 0 "
		case "w" '周
			wheresql = wheresql & " and datediff('ww',proupdatetime,'"&now()&"') = 0 "
		case else
			wheresql = wheresql & " and datediff('m',proupdatetime,'"&now()&"') = 0 "
		end select 
	end if
	'处理分页
	if isviewpage = "y" then
		pagesql = "select count(*) from products where proisrec = false " & wheresql
		recordcounts = eConn.execute(pagesql)(0)
		if recordcounts/pagesizes > int(recordcounts/pagesizes) then 
			pagecounts = int(recordcounts/pagesizes) + 1
		else
			pagecounts = int(recordcounts/pagesizes)
		end if
		if pagecounts < 1 then pagecounts = 1
	else
		intpage = 1
		pagecounts = 1	
	end if
	if ordertype = "置顶" then
		ordertype = " proistop asc,proupdatetime desc "
	elseif right(ordertype,2) = "点击" then
		ordertype = " proclicks desc,proupdatetime desc "
	elseif ordertype = "更新" then
		ordertype = " proupdatetime desc "
	elseif ordertype = "发布" then
		ordertype = " proaddtime desc "
	else
		ordertype = " proupdatetime desc "
	end if
	'处理SQL
	if intpage > 1 then
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (products a left join sorts b on a.prosort = b.sortid) where proid not in(select top "&(intpage-1)*pagesizes &" proid from products where proisrec = false "&wheresql&" order by "&ordertype&") and proisrec = false  "&wheresql&" order by "&ordertype&""
	else
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (products a left join sorts b on a.prosort = b.sortid) where proisrec = false  "&wheresql&"  order by "&ordertype&""
	end if

	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "<ul>[LIST]<li><img src=""{产品小图地址}"" /><a href=""{产品地址}"">{产品名称}</a><span更新时间}</span></li>[/LIST]</ul>"
	end if
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if

	'response.Write("--------"&tempTagStr&"----------")

	startStr = replace(startStr,"{栏目名称}",getSortName(sortid))
	endStr = replace(endStr,"{栏目名称}",getSortName(sortid))

	startStr = replace(startStr,"{栏目ID}",sortid)
	endStr = replace(endStr,"{栏目ID}",sortid)
	
	returnStr = startStr
	'response.Write(pagesql)
	eRs.open pagesql,eConn,0,1
	do while not eRs.eof
		'response.Write("aa")
		linenums = linenums + 1
		tempListStr = listStr
		tempPage = eRs("sortshowtemplates")
		listpage = eRs("sortlisttemplates")
		if tempPage = "" or isnull(tempPage) then tempPage = "proshow.asp"
		if listpage = "" or isnull(listpage) then listpage = "prolist.asp"
		if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
			if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
			else
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
			end if
			
		end if
		tempListStr = replace(tempListStr,"{i}",linenums)
		tempListStr = replace(tempListStr,"{产品名称}",getSubString(eRs("proname"),titlesize))
		tempListStr = replace(tempListStr,"{产品地址}",tempPage & "?proid=" & eRs("proid"))
		tempListStr = replace(tempListStr,"{产品ID}",eRs("proid"))
		tempListStr = replace(tempListStr,"{产品点击数}",eRs("proClicks"))
		tempListStr = replace(tempListStr,"{栏目名称}",inull(eRs("sortname")))
		tempListStr = replace(tempListStr,"{栏目地址}",listpage & "?sortid=" & eRs("prosort"))
		tempListStr = replace(tempListStr,"{栏目ID}",eRs("prosort"))
		tempListStr = replace(tempListStr,"{产品小图地址}",eRs("prosmallpic"))
		tempListStr = replace(tempListStr,"{产品大图地址}",eRs("probigpic"))
		tempListStr = replace(tempListStr,"{摘要}",getSubString(eRs("prodescription"),Abstractsize))		
		tempListStr = replace(tempListStr,"{更新时间}",formatTime(eRs("proupdatetime"),dateformat))
		'tempListStr = replace(tempListStr,"{uptime}",Format_Time(eRs("artupdatetime"),dateformat))
		'自定义字段 愁死我了
		tempListStr = RegExpDiyTag(tempListStr,eRs("prodiyvalue"))
		
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	eRs.Close
	returnStr = returnStr & endStr
	returnStr = unrepreg(returnStr)
	returnStr = replaceGlobal(returnStr)
	if isviewpage = "y" then
		if instr(returnStr,"{分页}") > 0 then
			returnStr = replace(returnStr,"{分页}",writepage(pagecounts,intpage))
		else
			returnStr = returnStr & writepage(pagecounts,intpage)
		end if 
	end if
	'response.Write(err.number)
	proList = returnStr
end function

'主打产品列表
function mainmm(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
	dim intpage,sortid,pagesizes,pagecounts,pagesql,recordcounts,wheresql,fidsql,linenums,keyword
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,tempPage,tempListStr,listpage
	pagesizes = listnums
	'处理栏目ID,如果有输入,直接赋值,没有,则获取sortid,如果还没有,则根据文件名获取,还是没有?那就全部吧..
	if sid <> 0 then
		sortid = sid
	else
		sortid = getSortID
	end if
	
	'处理分页
	intpage = request.QueryString("page")
	if intpage = "" then
		intpage = 1
	else 
		intpage = int(intpage)
	end if
	
	'处理条件语句
	if sortid <> 0 then
		wheresql = " and prosort = " & sortid
	end if
	if issubclass = "y" and sortid <> 0 then
		wheresql = " and prosort in (select sortid from sorts where sortid = "&sortid&" or sortid in (select sortid from sorts where sortfid="&sortid&" ) or sortfid in (select sortid from sorts where sortfid="&sortid&" )) "
	end if
	
	if ispiconly = "y" then
		wheresql = wheresql & " and prosmallpic <> '' "
	end if
	keyword = replace(request.QueryString("keyword"),"'","")
	if keyword <> "" then
		wheresql = wheresql & " and proname like '%"&keyword&"%' "
	end if
	
	if left(ordertype,2) = "点击" then
		select case lcase(right(ordertype,1))
		case "y" '年
			wheresql = wheresql & " and datediff('yyyy',proupdatetime,'"&now()&"') = 0 "
		case "j" '季
			wheresql = wheresql & " and datediff('q',proupdatetime,'"&now()&"') = 0 "
		case "m" '月
			wheresql = wheresql & " and datediff('m',proupdatetime,'"&now()&"') = 0 "
		case "w" '周
			wheresql = wheresql & " and datediff('ww',proupdatetime,'"&now()&"') = 0 "
		case else
			wheresql = wheresql & " and datediff('m',proupdatetime,'"&now()&"') = 0 "
		end select 
	end if
	'处理分页
	
	if isviewpage = "y" then
		pagesql = "select count(*) from mainProduct where proisrec = false " & wheresql
		recordcounts = eConn.execute(pagesql)(0)
		if recordcounts/pagesizes > int(recordcounts/pagesizes) then 
			pagecounts = int(recordcounts/pagesizes) + 1
		else
			pagecounts = int(recordcounts/pagesizes)
		end if
		if pagecounts < 1 then pagecounts = 1
	else
		intpage = 1
		pagecounts = 1	
	end if
	if ordertype = "置顶" then
		ordertype = " proistop asc,proupdatetime desc "
	elseif right(ordertype,2) = "点击" then
		ordertype = " proclicks desc,proupdatetime desc "
	elseif ordertype = "更新" then
		ordertype = " proupdatetime desc "
	elseif ordertype = "发布" then
		ordertype = " proaddtime desc "
	else
		ordertype = " proupdatetime desc "
	end if
	'处理SQL
	if intpage > 1 then
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (mainProduct a left join sorts b on a.prosort = b.sortid) where proid not in(select top "&(intpage-1)*pagesizes &" proid from mainProduct where proisrec = false "&wheresql&" order by "&ordertype&") and proisrec = false  "&wheresql&" order by "&ordertype&""
	else
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (mainProduct a left join sorts b on a.prosort = b.sortid) where proisrec = false  "&wheresql&"  order by "&ordertype&""
	end if
	
	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "<ul>[LIST]<li><img src=""{产品小图地址}"" /><a href=""{产品地址}"">{产品名称}</a><span更新时间}</span></li>[/LIST]</ul>"
	end if
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if

	'response.Write("--------"&tempTagStr&"----------")

	startStr = replace(startStr,"{栏目名称}",getSortName(sortid))
	endStr = replace(endStr,"{栏目名称}",getSortName(sortid))

	startStr = replace(startStr,"{栏目ID}",sortid)
	endStr = replace(endStr,"{栏目ID}",sortid)
	
	returnStr = startStr
	
	'response.Write(pagesql)
	eRs.open pagesql,eConn,0,1
	do while not eRs.eof
		
		'response.Write(eRs.recordcount)
		linenums = linenums + 1
		tempListStr = listStr
		tempPage = eRs("sortshowtemplates")
		listpage = eRs("sortlisttemplates")
		if tempPage = "" or isnull(tempPage) then tempPage = "mainProshow.asp"
		if listpage = "" or isnull(listpage) then listpage = "mainProlist.asp"
		if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
			if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
			else
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
			end if
			
		end if
		tempListStr = replace(tempListStr,"{i}",linenums)
		tempListStr = replace(tempListStr,"{产品名称}",getSubString(eRs("proname"),titlesize))
		tempListStr = replace(tempListStr,"{产品地址}",tempPage & "?mproid=" & eRs("proid"))
		tempListStr = replace(tempListStr,"{产品ID}",eRs("proid"))
		tempListStr = replace(tempListStr,"{产品点击数}",eRs("proClicks"))
		tempListStr = replace(tempListStr,"{栏目名称}",inull(eRs("sortname")))
		tempListStr = replace(tempListStr,"{栏目地址}",listpage & "?sortid=" & eRs("prosort"))
		tempListStr = replace(tempListStr,"{栏目ID}",eRs("prosort"))
		tempListStr = replace(tempListStr,"{产品小图地址}",eRs("prosmallpic"))
		tempListStr = replace(tempListStr,"{产品大图地址}",eRs("probigpic"))
		tempListStr = replace(tempListStr,"{摘要}",getSubString(eRs("prodescription"),Abstractsize))		
		tempListStr = replace(tempListStr,"{更新时间}",formatTime(eRs("proupdatetime"),dateformat))
		tempListStr = replace(tempListStr,"{价格}",eRs("proPrice"))
		tempListStr = replace(tempListStr,"{车辆}",eRs("proCar"))
		tempListStr = replace(tempListStr,"{机票}",eRs("proTicket"))
		tempListStr = replace(tempListStr,"{高尔夫}",getSubString(eRs("proGolf"),18))
		tempListStr = replace(tempListStr,"{餐饮}",getSubString(eRs("proEat"),9))
		tempListStr = replace(tempListStr,"{酒店}",getSubString(eRs("proHotel"),9))
		tempListStr = replace(tempListStr,"{其他}",getSubString(eRs("proElse"),18))
		tempListStr = replace(tempListStr,"{包含}",eRs("proIn"))
		tempListStr = replace(tempListStr,"{相关介绍}",eRs("prointro"))
		tempListStr = replace(tempListStr,"{注意事项}",eRs("proNote"))
		'tempListStr = replace(tempListStr,"{uptime}",Format_Time(eRs("artupdatetime"),dateformat))
		'自定义字段 愁死我了
		tempListStr = RegExpDiyTag(tempListStr,eRs("prodiyvalue"))
		
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	eRs.Close
	returnStr = returnStr & endStr
	returnStr = unrepreg(returnStr)
	returnStr = replaceGlobal(returnStr)
	if isviewpage = "y" then
		if instr(returnStr,"{分页}") > 0 then
			returnStr = replace(returnStr,"{分页}",writepage(pagecounts,intpage))
		else
			returnStr = returnStr & writepage(pagecounts,intpage)
		end if 
	end if
	
	'response.Write(err.number)
	mainmm = returnStr
end function



'酒店信息列表
function hotel(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
	dim intpage,sortid,pagesizes,pagecounts,pagesql,recordcounts,wheresql,fidsql,linenums,keyword
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,tempPage,tempListStr,listpage
	pagesizes = listnums
	'处理栏目ID,如果有输入,直接赋值,没有,则获取sortid,如果还没有,则根据文件名获取,还是没有?那就全部吧..
	if sid <> 0 then
		sortid = sid
	else
		sortid = getSortID
	end if
	
	'处理分页
	intpage = request.QueryString("page")
	if intpage = "" then
		intpage = 1
	else 
		intpage = int(intpage)
	end if
	
	'处理条件语句
	if sortid <> 0 then
		wheresql = " and prosort = " & sortid
	end if
	if issubclass = "y" and sortid <> 0 then
		wheresql = " and prosort in (select sortid from sorts where sortid = "&sortid&" or sortid in (select sortid from sorts where sortfid="&sortid&" ) or sortfid in (select sortid from sorts where sortfid="&sortid&" )) "
	end if
	
	if ispiconly = "y" then
		wheresql = wheresql & " and prosmallpic <> '' "
	end if
	keyword = replace(request.QueryString("keyword"),"'","")
	if keyword <> "" then
		wheresql = wheresql & " and proname like '%"&keyword&"%' "
	end if
	
	if left(ordertype,2) = "点击" then
		select case lcase(right(ordertype,1))
		case "y" '年
			wheresql = wheresql & " and datediff('yyyy',proupdatetime,'"&now()&"') = 0 "
		case "j" '季
			wheresql = wheresql & " and datediff('q',proupdatetime,'"&now()&"') = 0 "
		case "m" '月
			wheresql = wheresql & " and datediff('m',proupdatetime,'"&now()&"') = 0 "
		case "w" '周
			wheresql = wheresql & " and datediff('ww',proupdatetime,'"&now()&"') = 0 "
		case else
			wheresql = wheresql & " and datediff('m',proupdatetime,'"&now()&"') = 0 "
		end select 
	end if
	'处理分页
	
	if isviewpage = "y" then
		pagesql = "select count(*) from hotel where proisrec = false " & wheresql
		recordcounts = eConn.execute(pagesql)(0)
		if recordcounts/pagesizes > int(recordcounts/pagesizes) then 
			pagecounts = int(recordcounts/pagesizes) + 1
		else
			pagecounts = int(recordcounts/pagesizes)
		end if
		if pagecounts < 1 then pagecounts = 1
	else
		intpage = 1
		pagecounts = 1	
	end if
	if ordertype = "置顶" then
		ordertype = " proistop asc,proupdatetime desc "
	elseif right(ordertype,2) = "点击" then
		ordertype = " proclicks desc,proupdatetime desc "
	elseif ordertype = "更新" then
		ordertype = " proupdatetime desc "
	elseif ordertype = "发布" then
		ordertype = " proaddtime desc "
	else
		ordertype = " proupdatetime desc "
	end if
	'处理SQL
	if intpage > 1 then
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (hotel a left join sorts b on a.prosort = b.sortid) where proid not in(select top "&(intpage-1)*pagesizes &" proid from hotel where proisrec = false "&wheresql&" order by "&ordertype&") and proisrec = false  "&wheresql&" order by "&ordertype&""
	else
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (hotel a left join sorts b on a.prosort = b.sortid) where proisrec = false  "&wheresql&"  order by "&ordertype&""
	end if
	
	
	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "<ul>[LIST]<li><img src=""{产品小图地址}"" /><a href=""{产品地址}"">{产品名称}</a><span更新时间}</span></li>[/LIST]</ul>"
	end if
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if

	'response.Write("--------"&tempTagStr&"----------")

	startStr = replace(startStr,"{栏目名称}",getSortName(sortid))
	endStr = replace(endStr,"{栏目名称}",getSortName(sortid))

	startStr = replace(startStr,"{栏目ID}",sortid)
	endStr = replace(endStr,"{栏目ID}",sortid)
	
	returnStr = startStr
	
	'response.Write(pagesql)
	eRs.open pagesql,eConn,0,1
	do while not eRs.eof
		'response.Write("aa")
		linenums = linenums + 1
		tempListStr = listStr
		tempPage = eRs("sortshowtemplates")
		listpage = eRs("sortlisttemplates")
		if tempPage = "" or isnull(tempPage) then tempPage = "hotelShow.asp"
		if listpage = "" or isnull(listpage) then listpage = "hotel.asp"
		if RegExpTest(tempListStr,"{FH:\d:[^}]*}") then
			if linenums mod int(RegExpRet(tempListStr,"{FH:(\d):[^}]*}",0)) = 0 then
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),RegExpRet(tempListStr,"{FH:(\d):([^}]*)}",1))
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","$1")
			else
				'tempListStr = replace(tempListStr,RegExpRet(tempListStr,"({FH:\d:[^}]*})",0),"")
				tempListStr = ReplaceTest(tempListStr,"{FH:\d:([^}]*)}","")
			end if
			
		end if
		tempListStr = replace(tempListStr,"{i}",linenums)
		tempListStr = replace(tempListStr,"{产品名称}",getSubString(eRs("ProName"),titlesize))
		tempListStr = replace(tempListStr,"{产品地址}",tempPage & "?hproid=" & eRs("proid"))
		tempListStr = replace(tempListStr,"{产品ID}",eRs("proid"))
		tempListStr = replace(tempListStr,"{产品点击数}",eRs("proClicks"))
		tempListStr = replace(tempListStr,"{栏目名称}",inull(eRs("sortname")))
		tempListStr = replace(tempListStr,"{栏目地址}",listpage & "?sortid=" & eRs("prosort"))
		tempListStr = replace(tempListStr,"{栏目ID}",eRs("prosort"))
		tempListStr = replace(tempListStr,"{产品小图地址}",eRs("prosmallpic"))
		tempListStr = replace(tempListStr,"{产品大图地址}",eRs("probigpic"))
		tempListStr = replace(tempListStr,"{摘要}",getSubString(eRs("prodescription"),Abstractsize))		
		tempListStr = replace(tempListStr,"{更新时间}",formatTime(eRs("proupdatetime"),dateformat))
		tempListStr = replace(tempListStr,"{酒店星级}",eRs("proStar"))
		
		'tempListStr = replace(tempListStr,"{uptime}",Format_Time(eRs("artupdatetime"),dateformat))
		'自定义字段 愁死我了
		tempListStr = RegExpDiyTag(tempListStr,eRs("prodiyvalue"))
		
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	eRs.Close
	returnStr = returnStr & endStr
	returnStr = unrepreg(returnStr)
	returnStr = replaceGlobal(returnStr)
	if isviewpage = "y" then
		if instr(returnStr,"{分页}") > 0 then
			returnStr = replace(returnStr,"{分页}",writepage(pagecounts,intpage))
		else
			returnStr = returnStr & writepage(pagecounts,intpage)
		end if 
	end if
	
	'response.Write(err.number)
	hotel = returnStr
end function





' ----------------------
' 产品内容调用
' ----------------------
function proContent(style)
	Dim tempTagStr,returnStr,tempValue,intpage,breakpage,newsid,pagesplit,proid
	breakpage = "<div style=""page-break-after: always""><span style=""display: none"">&nbsp;</span></div>"
	intpage = request.QueryString("page")
	proid = request.QueryString("proid")
	if not isnumeric(intpage) then 
		intpage = 1
	else
		intpage = int(intpage)
	end if
	if not isnumeric(proid) then
		proContent = "无效ID"
		exit function
	else
		proid = int(proid)
	end if
	if intpage < 1 then intpage = 1
	if proid < 1 then exit function
	
	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "<h1>{产品名称}</h1><div id=""procontent"">{产品内容}</div>"
	end if
	
	'response.Write("select top 1 * from articles from artisrec = false and artid = " & newsid)
	eRs.open "select top 1 * from products where proisrec = false and proid = " & proid,eConn,0,1
	if eRs.eof then
		newsContent = "无效ID"
		exit function
	end if
	
	pagesplit = split(ers("procontent") & " ",breakpage)
	if intpage > ubound(pagesplit)+1 then intpage = ubound(pagesplit)+1
		tempTagStr = replace(tempTagStr,"{title}",eRs("proname"))
		tempTagStr = replace(tempTagStr,"{产品内容}",pagesplit(intpage-1))
		tempListStr = replace(tempListStr,"{产品名称}",getSubString(eRs("proname"),titlesize))
		tempListStr = replace(tempListStr,"{产品地址}",tempPage & "?proid=" & eRs("proid"))
		tempListStr = replace(tempListStr,"{产品ID}",eRs("proid"))
		tempListStr = replace(tempListStr,"{产品点击数}",eRs("proClicks"))
		tempListStr = replace(tempListStr,"{栏目名称}",inull(eRs("sortname")))
		tempListStr = replace(tempListStr,"{栏目地址}",listpage & "?sortid=" & eRs("prosort"))
		tempListStr = replace(tempListStr,"{产品小图地址}",eRs("prosmallpic"))
		tempListStr = replace(tempListStr,"{产品大图地址}",eRs("probigpic"))
		tempListStr = replace(tempListStr,"{摘要}",getSubString(eRs("prodescription"),Abstractsize))		
		tempListStr = replace(tempListStr,"{更新时间}",formatTime(eRs("proupdatetime"),dateformat))
		'tempListStr = replace(tempListStr,"{uptime}",Format_Time(eRs("artupdatetime"),dateformat))
		'自定义字段
		tempListStr = RegExpDiyTag(tempListStr,eRs("prodiyvalue"))
		
	eRs.Close
	returnStr = tempTagStr
	proContent = unrepreg(returnStr)
	returnStr = replaceGlobal(returnStr)
	if ubound(pagesplit) > 0 then
		if instr(returnStr,"{分页}") > 0 then
			returnStr = replace(returnStr,"{分页}",writepage(ubound(pagesplit)+1,intpage))
		else
			returnStr = returnStr & writepage(ubound(pagesplit)+1,intpage)
		end if 
	end if
	
	
	
end function

' ----------------------
' 当前导航位置
' ----------------------
function getLocalNav(style)
	dim sortid,newsid,proid,speakStr,aloneid
	Dim tempTagStr,returnStr,tempValue
	aloneid = request.QueryString("aloneid")
	sortid = request.QueryString("sortid")
	newsid = request.QueryString("newsid")
	proid = request.QueryString("proid")
	
	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "<a href=""default.asp"">首页</a> {nav:&gt;&gt;}"
	end if
	'解析全局标签  '如果当前位置解析不正确，注释下面此行
	tempTagStr = replaceGlobal(tempTagStr)
	' RegExpRet(tempTagStr,"([^\{]*)\{nav:([^\}]*)\}(.*)",0)
	speakStr = " " & RegExpRet(tempTagStr,"{nav:([^\}]*)\}",0) & " "
	if sortid <> "" then 
		getLocalNav = replace(tempTagStr,RegExpRet(tempTagStr,"({nav:[^\}]*\})",0),getSortNav(speakStr,int(sortid)))
		exit function
	end if
	if newsid <> "" then
		eRs.open "select top 1 arttitle,artid,artsort from articles where artid = " & int(newsid),eConn,0,1
		if not eRs.eof  then
			'tempTagStr = tempTagStr & speakStr & eRs("arttitle")
			tempTagStr = replace(tempTagStr,RegExpRet(tempTagStr,"({nav:[^\}]*\})",0),getSortNav(speakStr,int(eRs("artsort")))& speakStr & eRs("arttitle") )
			
		end if
		eRs.Close
		getLocalNav = tempTagStr
		exit function
	end if
	if proid <> "" then
		eRs.open "select top 1 proid,proname,prosort from products where proid = " & int(proid),eConn,0,1
		if not eRs.eof  then
			tempTagStr = replace(tempTagStr,RegExpRet(tempTagStr,"({nav:[^\}]*\})",0),getSortNav(speakStr,int(eRs("prosort"))) &  speakStr & eRs("proname") )
			'tempTagStr = tempTagStr & speakStr & eRs("proname")
		end if
		eRs.Close
		getLocalNav = tempTagStr
		exit function
	end if
		'response.Write(T("sid"))
	if T("sid") <> 0 then
		getLocalNav = replace(tempTagStr,RegExpRet(tempTagStr,"({nav:[^\}]*\})",0),getSortNav(speakStr,T("sid")))
		exit function
	end if
	'response.Write(tSortID)
	if tSortID <> 0 then
		getLocalNav = replace(tempTagStr,RegExpRet(tempTagStr,"({nav:[^\}]*\})",0),getSortNav(speakStr,tSortID))
		exit function
	end if
	
	getLocalNav = ""
end function

' ----------------------
' 一个栏目的完整路径
' sid 栏目ID~
' speakStr 分隔符,默认为 >> 
' ----------------------
function getSortNav(speakStr,sid)
	dim sortid
	if sid = 0 then
		sortid = request.QueryString("sortid")
		if sortid = "" then 
			on error resume next
			sortid = eConn.execute("select top 1 sortid from sorts where sortlisttemplates = '"&getFileName()&"'" )(0)
			if err.number <> 0 then
				err.clear
				getSortNav = ""
				exit function
			end if
		end if
	else
		sortid = int(sid)
	end if
	
	sortStr = ""
	if speakStr = "" then
		speakStr = "&gt;&gt;"
	end if
	
	getSortFidNav speakStr,sortid
	getSortNav = sortStr
	
end function

' ----------------------
' 这玩印儿不是用来调用的
' ----------------------
sub getSortFidNav(speakStr,sid)
	dim sortRs,templist
	
	set sortRs = eConn.execute("select top 1 sortid,sortname,sortfid,sortlisttemplates from sorts where sortid = "&sid)
	if not sortRs.eof then
		templist = sortRs("sortlisttemplates")
		if templist = "" then 
			templist = "newslist.asp"
		end if
		sortStr = speakStr & " <a href="""&templist & "?sortid=" & sortRs(0)&""">"&sortRs(1)&"</a> "  & sortStr
		if int(sortRs(2)) <> 0 then
			getSortFidNav speakStr ,sortRs(2)
		end if
	end if
	sortRs.Close
	set sortRs = nothing
end sub

' ----------------------
' 留言框
' ----------------------
function addMsg(style)
	Dim tempTagStr,returnStr,tempValue,msgJs,sortid
	sortid = request.QueryString("sortid")
	if sortid = "" then 
			on error resume next
			sortid = eConn.execute("select top 1 sortid from sorts where sortlisttemplates = '"&getFileName()&"'" )(0)
			if err.number <> 0 then
				err.clear
				sortid = ""
			end if
	end if
		
	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then  '默认表单
		tempTagStr = "<form action="""" method=""post"" onsubmit=""return false;"">" & vbcrlf &_
					"				  <div class=""msg_form"">" & vbcrlf &_
					"					<div class=""form_row"">" & vbcrlf &_
					"					  <label class=""contact"">用户名:</label>" & vbcrlf &_
					"					  <input type=""text"" name=""username"" id=""username"" /><span id=""userTip""></span>" & vbcrlf &_
					"					</div>" & vbcrlf &_
					"					<div class=""form_row"">" & vbcrlf &_
					"					  <label class=""contact"">Email:</label>" & vbcrlf &_
					"					  <input type=""text"" name=""email"" id=""email"" value="""" /><span id=""emailTip""></span>" & vbcrlf &_
					"					</div>" & vbcrlf &_
					"					<div class=""form_row"">" & vbcrlf &_
					"					  <label class=""contact"">电话:</label>" & vbcrlf &_
					"					  <input type=""text"" name=""phone"" id=""phone"" /><span id=""phoneTip""></span>" & vbcrlf &_
					"					</div>" & vbcrlf &_
					"					<div class=""form_row"">" & vbcrlf &_
					"					  <label class=""contact"">内容:</label>" & vbcrlf &_
					"					  <textarea rows=""10"" cols=""40"" name=""content"" id=""content"" ></textarea><span id=""contentTip""></span>" & vbcrlf &_
					"					</div>" & vbcrlf &_
					"					<div id=""stateStr""></div>" & vbcrlf &_
					"					<div class=""form_row"">" & vbcrlf &_
					"					  <input type=""image"" onclick=""msgsubmit()"" src=""images/send.gif"" class=""send""/>" & vbcrlf &_
					"					</div>" & vbcrlf &_
					"				  </div>" & vbcrlf &_
					"				</div>" & vbcrlf &_
					"			</form>"
	end if
	'解析全局标签
	tempTagStr = replaceGlobal(tempTagStr)
	msgJs = "<script language=""javascript"" type=""text/javascript"" src=""script/jquery132.js""></script>" & vbcrlf &_
			"<script language=""javascript"" type=""text/javascript"">" & vbcrlf &_
			"function msgsubmit(){" & vbcrlf &_
			"	if ($('#username').val()==''){$('#userTip').html(""用户名不能为空"");return;}else{$('#userTip').html("""");}" & vbcrlf &_
			"	if ($('#content').val()==''){$('#contentTip').html(""内容不能为空"");return;}else{$('#contentTip').html("""");}" & vbcrlf &_
			"	if ($('#email').val()!='' && !/^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$/.test($('#email').val())){{$('#emailTip').html('Email格式错误');return;}}else{$('#emailTip').html('');}" & vbcrlf &_
			"	$('#stateStr').html(""正在提交..."");" & vbcrlf &_
			"	$('#stateStr').show();" & vbcrlf &_
			"	$.ajax({type:'post'," & vbcrlf &_
			"		   url:'include/saveMsg.asp'," & vbcrlf &_
			"		   data:'act=msg&sortid="&sortid&"&username='+escape($('#username').val())+'&email='+escape($('#email').val())+'&phone='+escape($('#phone').val())+'&content='+escape($('#content').val())," & vbcrlf &_
			"		   success:function(msg){" & vbcrlf &_
			"			   if (msg==""0""){$('#stateStr').html('提交成功,请等待审核或回复处理!');}" & vbcrlf &_
			"			   else {$('#stateStr').html('提交失败:'+msg);}" & vbcrlf &_
			"			}" & vbcrlf &_
			"		   });" & vbcrlf &_
			"}" & vbcrlf &_
			"</script>"

	tempTagStr = msgJs & tempTagStr
	session("ismsg") = "esmsg"
	addMsg = tempTagStr
end function

		
' ----------------------
' 留言列表函数
' sid 为栏目ID
' style 样式名称
' listnums 调用条数,如果是终级列表,则为每页数量
' dateformat 日期格式,
' contentsize 内容显示长度 中文一个算俩个,英文一个算一个
' isviewpage 是否显示分页 1 显示 0 不显示
' ----------------------
function msglist(sid,style,listnums,dateformat,contentsize,isviewpage)
	dim intpage,sortid,pagesizes,pagecounts,pagesql,recordcounts,wheresql,fidsql,linenums
	Dim tempTagStr,returnStr,tempValue,listStr,startStr,endStr,tempPage,tempListStr,listpage
	pagesizes = listnums
	'处理栏目ID,如果有输入,直接赋值,没有,则获取sortid,如果还没有,则根据文件名获取,还是没有?那就全部吧..
	if sid <> 0 then
		sortid = sid
	else
		sortid = request.QueryString("sortid")
		if sortid  = "" then 
			on error resume next
			sortid = eConn.execute("select top 1 sortid from sorts where sortlisttemplates = '"&getFileName()&"'" )(0)
			if err.number <> 0 then
				err.clear
				sortid = 0
			end if
		else 
			sortid = int(sortid)
		end if
	end if
	'处理分页
	intpage = request.QueryString("page")
	if intpage = "" then
		intpage = 1
	else 
		intpage = int(intpage)
	end if
	'处理条件语句
	if sortid <> 0 then
		wheresql = " and msgsort = " & sortid
	end if

	'处理分页
	if isviewpage = "y" then
		pagesql = "select count(*) from messages where msgisaudit = true " & wheresql
		recordcounts = eConn.execute(pagesql)(0)
		if recordcounts/pagesizes > int(recordcounts/pagesizes) then 
			pagecounts = int(recordcounts/pagesizes) + 1
		else
			pagecounts = int(recordcounts/pagesizes)
		end if
		if pagecounts < 1 then pagecounts = 1
	else
		intpage = 1
		pagecounts = 1	
	end if
	'处理SQL
	if intpage > 1 then
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (messages a left join sorts b on a.msgsort = b.sortid) where msgid not in(select top "&(intpage-1)*pagesizes &" msgid from messages where msgisaudit = true "&wheresql&" order by msgaddtime desc) and msgisaudit = true  "&wheresql&" order by msgaddtime desc"
	else
		pagesql = "select top "&pagesizes&" a.*,b.sortshowtemplates,b.sortlisttemplates,b.sortname from (messages a left join sorts b on a.msgsort = b.sortid) where msgisaudit = true  "&wheresql&"  order by msgaddtime desc"
	end if

	tempTagStr = getTagStr(style)
	
	if tempTagStr = "" or isnull(tempTagStr) then
		tempTagStr = "[LIST]<div class=""msglist"">" & vbcrlf &_
				"      <div class=""line""><lable>用户名:</lable>{用户名}</div>" & vbcrlf &_
				"      <div class=""content""><lable>留言:</lable>{留言内容}</div>" & vbcrlf &_
				"</div>" & vbcrlf &_
				"[RE]<div class=""msgre"">{回复内容}</div>[/RE][/LIST]"
	end if
	tempTagStr = replaceGlobal(tempTagStr)
	if instr(tempTagStr,"[LIST]") > 0 and instr(tempTagStr,"[/LIST]") then
		startStr = mid(tempTagStr,1,instr(tempTagStr,"[LIST]")-1)
		listStr = mid(tempTagStr,instr(tempTagStr,"[LIST]")+6,instr(tempTagStr,"[/LIST]") - instr(tempTagStr,"[LIST]") - 6)
		endStr = mid(tempTagStr,instr(tempTagStr,"[/LIST]")+7)
	else
		startStr = ""
		endStr = ""
		listStr = tempTagStr
	end if

	'response.Write("--------"&tempTagStr&"----------")
	startStr = replace(startStr,"{栏目名称}",getSortName(sortid))
	endStr = replace(endStr,"{栏目名称}",getSortName(sortid))

	startStr = replace(startStr,"{栏目ID}",sortid)
	endStr = replace(endStr,"{栏目ID}",sortid)
	
	returnStr = startStr
	eRs.open pagesql,eConn,0,1
	do while not eRs.eof
		linenums = linenums + 1
		tempListStr = listStr
		if instr(tempListStr,"[RE]") > 0 then
			if inull(eRs("msgreply")) <> "" then
				tempListStr = replace(tempListStr,"{回复内容}",eRs("msgreply"))
				tempListStr = replace(tempListStr,"[RE]","")
				tempListStr = replace(tempListStr,"[/RE]","")
			else
				tempListStr = replaceTest(tempListStr,"\[RE\][^\[]+\[\/RE\]","")
			end if
			
		end if
		tempListStr = replace(tempListStr,"{i}",linenums)
		tempListStr = replace(tempListStr,"{栏目名称}",inull(eRs("sortname")))
		tempListStr = replace(tempListStr,"{栏目地址}",listpage & "?sortid=" & eRs("msgsort"))
		tempListStr = replace(tempListStr,"{栏目ID}",eRs("msgsort"))
		tempListStr = replace(tempListStr,"{用户名}",eRs("msgname"))
		tempListStr = replace(tempListStr,"{电话}",eRs("msgtel"))
		tempListStr = replace(tempListStr,"{email}",eRs("msgemail"))
		tempListStr = replace(tempListStr,"{留言内容}",getSubString(eRs("msgcontent"),contentsize))
		tempListStr = replace(tempListStr,"{留言IP}",eRs("msgip"))		
		tempListStr = replace(tempListStr,"{留言时间}",FormatTime(eRs("msgaddtime"),dateformat))
		'tempListStr = replace(tempListStr,"{uptime}",Format_Time(eRs("artupdatetime"),dateformat))
		returnStr = returnStr & tempListStr
		eRs.movenext
	loop
	eRs.Close
	returnStr = returnStr & endStr
	'解析全局标签
	returnStr = unrepreg(returnStr)
	returnStr = replaceGlobal(returnStr)
	if isviewpage = "y" then
		if instr(returnStr,"{分页}") > 0 then
			returnStr = replace(returnStr,"{分页}",writepage(pagecounts,intpage))
		else
			returnStr = returnStr & writepage(pagecounts,intpage)
		end if 
	end if
	msglist = returnStr
end function




' -------------------------
' 默认要载入的数组
' --------------------------
sub Init()
	dim nRs,fi
	set nRs = server.CreateObject("adodb.recordset")
	if request.QueryString("sortid") <> "" then
		'表示在栏目页或单页 ?
		nRs.open "select top 1 sortid,SortName,sorttype,SortKeywords,SortDescription,SortPic,SortContent from sorts where sortid = " & int(request.QueryString("sortid")),eConn,0,1
		if not nRs.eof then
			tSortType = "sort"
			tSortID = nRs("sortid")
			for fi = 1 to 7
				fieldArray(fi) = nRs(fi-1)
			next
			if int(nRs("sorttype")) = 1 then
				nRs.close
				nRs.open "select top 1 AloneID,AloneName,AlonePic,AloneSort,AloneContent,AloneAbstract,(select top 1 sortname from sorts where sortid = a.alonesort) as sortname from alones a where aloneissort = " & int(request.QueryString("sortid")),eConn,0,1
				if not nRs.eof then
					for fi = 11 to 17
						fieldArray(fi) = nRs(fi-11)
					next
				end if
			end if
		end if
		nRs.Close
	elseif request.QueryString("aloneid") <> "" then
		nRs.open "select top 1 AloneID,AloneName,AlonePic,AloneSort,AloneContent,AloneAbstract,(select top 1 sortname from sorts where sortid = a.alonesort) as sortname from alones a where aloneid = " & int(request.QueryString("aloneid")),eConn,0,1
		  if not nRs.eof then
			  for fi = 11 to 17
				  fieldArray(fi) = nRs(fi-11)
			  next
			  tSortType = "alone"
			  if eint(fieldArray(17)) <> 0 then
			  	tSortID = fieldArray(17) '?
			  else
			  	tSortID = fieldArray(14)
			  end if
			  'response.Write("aa")
		  end if
		nRs.Close
	elseif request.QueryString("newsid") <> "" then
		nRs.open "select top 1 artid,arttitle,artclicks,artupdatetime,artsort,artsource,artauthor,artkeywords,artdescription,artcontent,(select top 1 sortname from sorts where sortid = a.artsort) as sortname,artpic from articles a where artid = " & int(request.QueryString("newsid")),eConn,0,1
		if not nRs.eof then
			for fi = 21 to 32
				fieldArray(fi) = nRs(fi-21)
			next
			tSortType = "news"
			tSortID = fieldArray(25)
			econn.execute("update articles set artclicks = artclicks + 1 where artid = " & nRs(0))
		end if
		nRs.Close
	elseif request.QueryString("proid") <> "" then
		nRs.open "select top 1 proid,proname,prosort,prosmallpic,probigpic,proupdatetime,prokeywords,prodescription,procontent,(select top 1 sortname from sorts where sortid = a.prosort) as sortname,proDiyValue from products a where proid = " & int(request.QueryString("proid")),eConn,0,1
		if not nRs.eof then
			for fi = 41 to 51
				fieldArray(fi) = nRs(fi-41)
			next
			diyfield = fieldArray(51)
			tSortType = "pros"
			tSortID = fieldArray(43)
			eConn.execute("update products set proClicks = proClicks + 1 where proid = " & nRs(0))
		end if
		nRs.Close
	elseif request.QueryString("mproid") <> "" then

		nRs.open "select top 1 proid,proname,prosort,prosmallpic,probigpic,proupdatetime,prokeywords,prodescription,procontent,(select top 1 sortname from sorts where sortid = a.prosort) as sortname,proDiyValue,proPrice,proCar,proTicket,proGolf,proEat,proHotel,proElse,proIn,prointro,proNote from mainProduct a where proid = " & int(request.QueryString("mproid")),eConn,0,1
		if not nRs.eof then
			for fi = 41 to 61
				fieldArray(fi) = nRs(fi-41)
			next
			diyfield = fieldArray(51)
			tSortType = "pros"
			tSortID = fieldArray(43)
			eConn.execute("update mainProduct set proClicks = proClicks + 1 where proid = " & nRs(0))
		end if
		nRs.Close
		
	elseif request.QueryString("hproid") <> "" then

		nRs.open "select top 1 proid,proname,prosort,prosmallpic,probigpic,proupdatetime,prokeywords,prodescription,procontent,(select top 1 sortname from sorts where sortid = a.prosort) as sortname,proDiyValue,proStar,proPhone,proZone,proSub from hotel a where proid = " & int(request.QueryString("hproid")),eConn,0,1
		if not nRs.eof then
			for fi = 41 to 55
				fieldArray(fi) = nRs(fi-41)
			next
			
		
			
			eConn.execute("update hotel set proClicks = proClicks + 1 where proid = " & nRs(0))
		end if
		nRs.Close
		
		
	end if
	
	
	
	set nRs = nothing
end sub

' ---------------------------
' 原F万能函数 ' 此函数已废除,用T代替
' ---------------------------
function f(fieldname,pres) 
	f = t(fieldname & "$" & pres)
	exit function
end function




' ---------------------------
' 解析全局标签,这个有些问题
' ---------------------------
function replaceGlobal(templateStr)
	on error resume next
	  Dim regEx, Match, Matches,RetStr
	  Set regEx = New RegExp
	  regEx.Pattern = "{([^}]+)}"
	  regEx.IgnoreCase = False 
	  regEx.Global = True
	  Set Matches = regEx.Execute(templateStr)
	  RetStr = templateStr
	  For Each Match in Matches  '标签会引起死循环吗？
		RetStr = replace(RetStr,Match.Value,repreg(T(Match.SubMatches(0))))  '为防止循环只能套一次
	  Next
	replaceGlobal = unrepreg(RetStr)
	if err.number <> 0 then
		err.clear
		replaceGlobal = templateStr
	end if
end function

' ----------------------
' 简化的标签调用 改进的F标签 只有一个参数 格式 T("标签$参数1,参数2...")
' ----------------------
function T(tagParameterStr)
	on error resume next
	dim tagName,tagSplit,tagParSplit,tagParNums,tagIndex
	dim returnStr,tempRetStr
	dim tempWH
	dim tagParArray(40) '最多40参数?
	dim breakpage,intpage
	dim pagesplit,pagecontent
	
	intpage = request.QueryString("page")
	if not isnumeric(intpage) then 
		intpage = 1
	else
		intpage = int(intpage)
	end if	
	breakpage = "<div style=""page-break-after: always""><span style=""display: none"">&nbsp;</span></div>"
	
	if tagParameterStr = "" then exit function
	
	tagSplit = split(tagParameterStr,"$")
	if ubound(tagSplit) > 0 then
		tagParSplit = split(tagSplit(1),",")
		tagName = tagSplit(0)
		tagParNums = ubound(tagParSplit) + 1
	else
		tagParSplit = array(0)
		tagName = tagSplit(0)
		tagParNums = 1
	end if
	for tagIndex = 1 to tagParNums
		tagParArray(tagIndex) = tagParSplit(tagIndex-1)
	next	

	'response.Write(getlenstr(Application("web_name"),tagParSplit(0)) & err.number)
	select case lcase(tagName)
		'--------网站系统
		case "webname","wn","网站名称"
			returnstr = getlenstr(Application("web_name"),eint(tagParArray(1)))
		case "webcopyright","wc","网站版权"
			returnstr = getlenstr(Application("web_copyright"),eint(tagParArray(1)))
		case "webhost","wh","网站域名"	
			returnstr = getlenstr(Application("web_host"),eint(tagParArray(1)))
		case "webkeywords","wk","网站关键词"
			returnstr = getlenstr(Application("web_keywords"),eint(tagParArray(1)))
		case "webdescription","wd","网站说明"
			returnstr = getlenstr(Application("web_description"),eint(tagParSplit(0)))
		case "webbeian","wb","网站备案号"
			returnstr = Application("web_beian")
		'--------栏目 sortid,SortName,sorttype,SortKeywords,SortDescription,SortPic,SortContent
		case "sortid","sid","栏目ID"
			if fieldArray(1) <> "" and eint(fieldArray(1)) <> 0 then
				returnstr = fieldArray(1)
			elseif fieldArray(14) <> "" and eint(fieldArray(14)) <> 0 then
				returnstr = fieldArray(14)
			elseif fieldArray(25) <> "" and eint(fieldArray(25)) <> 0 then
				returnstr = fieldArray(25)
			elseif fieldArray(43) <> "" and eint(fieldArray(43)) <> 0 then
				returnstr = fieldArray(43)
			else
				returnstr = getSortID
			end if
		case "sortname","sname","stitle","st","栏目标题","栏目名称"
			if fieldArray(2) <> "" then
				returnstr = getlenstr(fieldArray(2),eint(tagParArray(1)))
			elseif fieldArray(17) <> "" then
				returnstr = getlenstr(fieldArray(17),eint(tagParArray(1)))
			elseif fieldArray(31) <> "" then
				returnstr = getlenstr(fieldArray(31),eint(tagParArray(1)))
			elseif fieldArray(50) <> "" then
				returnstr = getlenstr(fieldArray(50),eint(tagParArray(1)))
			elseif request.QueryString("keyword") <> "" then
				returnstr = "搜索:" & RemoveHTML(request.QueryString("keyword"))
			else
				returnstr = ""
			end if
		case "sortkeywords","skey","sk","栏目关键词"
			returnstr = getlenstr(fieldArray(4),eint(tagParArray(1)))
		case "提示信息"
			
			returnstr = eConn.execute("select "&tagSplit(1)&" from Tips where ID = 2")(tagSplit(1))
			'returnstr = tagSplit(1)
		case "Tips2"
			returnstr = Tips(2)
		case "sortdescription","sdesc","sd","栏目说明"
			returnstr = getlenstr(fieldArray(5),eint(tagParArray(1)))
		case "sortpic","sp","栏目图片"
			tempWH = ""
			if eint(tagParArray(1)) = 0 then
				tempRetStr = fieldArray(6)
			else
				tempRetStr = eConn.execute("select sortpic from sorts where sortid = " & eint(tagParArray(1)))(0)
			end if
			if inull(tagParArray(2)) <> "" then
				tempWH = " width="""& tagParArray(2) &""" "
			end if
			
			if inull(tagParArray(3)) <> "" then
				tempWH = tempWH & " height="""&tagParArray(3)&""" "
			end if 				
			if inull(tagParArray(4)) <> "" then
				tempWH = tempWH & " class="""&tagParArray(4)&""" "
			end if 	
			if inull(tagParArray(5)) <> "" then
				tempWH = tempWH  & " " & tagParArray(5) & " "
			end if 
			if inull(tempRetStr) <> "" then
				returnstr = "<img src=""" & tempRetStr & """ "  & tempWH &" border=""0"" />"
			else
				returnstr = ""
			end if
			'returnstr = fieldArray(5)
		case "sortcontent","sc","栏目简介"
			returnstr = getlenstr(fieldArray(7),eint(tagParArray(1)))
		'------------单页11 AloneID,AloneName,AlonePic,AloneSort,AloneContent,AloneAbstract
		case "aid","aboutid","aloneid","单页ID"
			returnstr = fieldArray(11)
		case "atitle","at","aname","单页名称"
			'response.Write(fieldArray(12))
			returnstr = getlenstr(fieldArray(12),eint(tagParArray(1)))
		case "apic","aboutpic","alonepic","ap","单页图片"   '加一参数ID~ T("单页图片,ID,width,height,class,pre")  最后一个参数是输出任意字符串,用来调用如onclick之类的东西
			tempWH = ""
			if eint(tagParArray(1)) = 0 then
				tempRetStr = fieldArray(13)
			else
				tempRetStr = eConn.execute("select AlonePic from alones where aloneID = " & eint(tagParArray(1)))(0)
			end if
			if inull(tagParArray(2)) <> "" then
				tempWH = " width="""& tagParArray(2) &""" "
			end if
			
			if inull(tagParArray(3)) <> "" then
				tempWH = tempWH & " height="""&tagParArray(3)&""" "
			end if
			if inull(tagParArray(4)) <> "" then
				tempWH = tempWH & " class="""&tagParArray(4)&""" "
			end if
			if inull(tagParArray(5)) <> "" then
				tempWH = tempWH  & " " & tagParArray(5) & " "
			end if 
			if inull(tempRetStr) <> "" then
				returnstr = "<img src=""" & tempRetStr & """ "  & tempWH &" border=""0"" />"
			else
				returnstr = ""
			end if
		case "picurl","单页图片地址"
			if eint(tagParArray(1)) = 0 then
				tempRetStr = fieldArray(13)
			else
				tempRetStr = eConn.execute("select AlonePic from alones where aloneID = " & eint(tagParArray(1)))(0)
			end if
			returnstr = tempRetStr
		case "aabstract","aa","aboutabstract","单页摘要"   '参数 id,len,ishtml
			if eint(tagParArray(1)) = 0 then
				tempRetStr = fieldArray(16)
			else
				tempRetStr = eConn.execute("select AloneAbstract from alones where aloneID = " & eint(tagParArray(1)))(0)
			end if
			if inull(tagParArray(3)) <> "html" then
				tempRetStr = RemoveHTML(tempRetStr)
			end if
			returnstr = getlenstr(tempRetStr,eint(tagParArray(2)))
		case "acontent","ac","aboutcontent","单页内容"
			if eint(tagParArray(1)) = 0 then
				tempRetStr = fieldArray(15)
			else
				tempRetStr = eConn.execute("select AloneContent from alones where aloneID = " & eint(tagParArray(1)))(0)
			end if
			if inull(tagParArray(3)) = "nohtml" then
				tempRetStr = RemoveHTML(tempRetStr)
			end if
			'returnstr =  eint(tagParArray(1))
			returnstr = getlenstr(tempRetStr,eint(tagParArray(2)))
			
		'-----------新闻内页21 artid,arttitle,artclicks,artupdatetime,artsort,artsource,artauthor,artkeywords,artdescription,artcontent,artpic
		case "nid","artid","newsid","文章ID","新闻ID"
			returnstr = fieldArray(21)
		case "ntitle","nt","newstitle","文章标题","新闻标题"
			returnstr = getlenstr(fieldArray(22),eint(tagParArray(1)))
		case "ntime","nuptime","arttime","nm","文章更新时间","新闻更新时间"
			returnstr = formatTime(fieldArray(24),tagParArray(1))
		case "nsource","ns","newssource","新闻来源","文章来源"
			returnstr = getlenstr(fieldArray(26),eint(tagParArray(1)))
		case "nauthor","newsauthor","na","文章作者","新闻作者"
			returnstr = getlenstr(fieldArray(27),eint(tagParArray(1)))
		case "nabstract","newsabstract","ne","ndescription","nd","文章说明","文章摘要","新闻摘要"
			returnstr = getlenstr(fieldArray(29),eint(tagParArray(1)))
		case "nkeywords","nk","newskeywords","文章关键词","新闻关键词"
			returnstr = getlenstr(fieldArray(28),eint(tagParArray(1)))
		case "ncontent","nc","newscontent","文章内容","新闻内容"
			if eint(tagParArray(1)) = 0 then  '长度为全部时,要分页

				pagesplit = split(fieldArray(30) & " ",breakpage)
				if intpage > ubound(pagesplit)+1 then intpage = ubound(pagesplit)+1
				'response.Write(intpage)
				if intpage < 1 then intpage =1
				if ubound(pagesplit) > 0 then
					returnstr = pagesplit(intpage-1) & writepage(ubound(pagesplit)+1,intpage)
				else
					returnstr = fieldArray(30)
				end if
	
				'returnstr = getlenstr(fieldArray(30),eint(tagParArray(1)))
			else
				returnstr = getlenstr(fieldArray(30),eint(tagParArray(1)))
			end if
			returnstr = unrepreg(returnstr)
			'------------------------------------------
		case "上一篇新闻","prevnews","上一篇文章","上一文章"
			if eint(fieldArray(21)) = 0 then
				returnstr = "暂无上一篇文章"
			else
				'response.write "select top 1 artid,arttitle from articles where artid < " & eint(fieldArray(21)) & " and sortid = " & T("sid")
				eRs.open "select top 1 artid,arttitle from articles where artid < " & eint(fieldArray(21)) & " and artsort = " & T("sid"),eConn,0,1
				if eRs.eof then
					returnstr = "暂无上一篇文章"
				else
					returnstr = "<a href=""?newsid="&eRs(0)&""">"&getlenstr(eRs(1),eint(tagParArray(1)))&"</a>"
				end if
				eRs.Close
			end if
			
		case "下一篇新闻","prevnews","上一篇文章","下一文章"
			if eint(fieldArray(21)) = 0 then
				returnstr = "暂无下一篇文章"
			else
				eRs.open "select top 1 artid,arttitle from articles where artid > " & eint(fieldArray(21)) & " and artsort = " & T("sid"),eConn,0,1
				if eRs.eof then
					returnstr = "暂无下一篇文章"
				else
					returnstr = "<a href=""?newsid="&eRs(0)&""">"&getlenstr(eRs(1),eint(tagParArray(1)))&"</a>"
				end if
				eRs.Close
			end if
		case "npic","newspic","新闻图片"   'npic$width,height,class,diy
			tempWH = ""
			'response.Write(inull(tagParArray(1)))
			if eint(tagParArray(1)) <> 0 then
				tempWH = tempWH & " width="""&tagParArray(1)&""" "
			end if 	
			if eint(tagParArray(2)) <> 0 then
				tempWH = tempWH & " height="""&tagParArray(2)&""" "
			end if 				
			if inull(tagParArray(3)) <> "" then
				tempWH = tempWH & " class="""&tagParArray(3)&""" "
			end if 	
			if inull(tagParArray(4)) <> "" then
				tempWH = tempWH  & " " & tagParArray(4) & " "
			end if 
			
			returnstr = "<img src=""" & fieldArray(32) & """ "  & tempWH &" border=""0"" />"
		case "newspicurl","新闻图片地址"
			returnstr = fieldArray(32)
			
		'产品内页41 proid,proname,prosort,prosmallpic,probigpic,proupdatetime,prokeywords,prodescription,procontent
		case "proid","pid","产品ID"
			returnstr = fieldArray(41)
		case "proname","pname","pn","产品名称"
			returnstr = getlenstr(fieldArray(42),eint(tagParArray(1)))
		
		case "prosmallpic","pspic","ps","产品小图"
			tempWH = ""
			if eint(tagParArray(1)) <> 0 then
				tempWH = tempWH & " width="""&tagParArray(1)&""" "
			end if 	
			if eint(tagParArray(2)) <> 0 then
				tempWH = tempWH & " height="""&tagParArray(2)&""" "
			end if 				
			if inull(tagParArray(3)) <> "" then
				tempWH = tempWH & " class="""&tagParArray(3)&""" "
			end if 	
			if inull(tagParArray(4)) <> "" then
				tempWH = tempWH  & " " & tagParArray(4) & " "
			end if 
			returnstr = "<img src=""" & fieldArray(44) & """ "  & tempWH &" border=""0"" />"
			
		case "产品小图地址"
			returnstr = fieldArray(44)
		case "probigpic","pbpic","pb","产品大图"
			tempWH = ""
			if inull(tagParArray(1)) <> "" then
				tempWH = tempWH & " width="""&tagParArray(2)&""" "
			end if 	
			if inull(tagParArray(2)) <> "" then
				tempWH = tempWH & " height="""&tagParArray(2)&""" "
			end if 				
			if inull(tagParArray(3)) <> "" then
				tempWH = tempWH & " class="""&tagParArray(3)&""" "
			end if 	
			if inull(tagParArray(4)) <> "" then
				tempWH = tempWH  & " " & tagParArray(4) & " "
			end if 
			returnstr = "<img src=""" & fieldArray(45) & """ "  & tempWH &" border=""0"" />"
		case "产品大图地址"
			returnstr = fieldArray(45)
		case "proupdatetime","ptime","pm","产品更新时间"
			returnstr = formatTime(fieldArray(46),tagParArray(1))
		case "prokeywords","pk","产品关键词"
			returnstr = getlenstr(fieldArray(47),eint(tagParArray(1)))
		case "prodescription","pr","产品说明"
			returnstr = getlenstr(fieldArray(48),eint(tagParArray(1)))
		case "上一篇产品","prevpro","上一个产品","上一产品"
			if eint(fieldArray(41)) = 0 then
				returnstr = "暂无上一个产品"
			else
				'response.write "select top 1 artid,arttitle from articles where artid < " & eint(fieldArray(21)) & " and sortid = " & T("sid")
				eRs.open "select top 1 proid,proname from products where proid < " & eint(fieldArray(41)) & " and prosort = " & T("sid"),eConn,0,1
				if eRs.eof then
					returnstr = "暂无上一个产品"
				else
					returnstr = "<a href=""?proid="&eRs(0)&""">"&getlenstr(eRs(1),eint(tagParArray(1)))&"</a>"
				end if
				eRs.Close
			end if
			
		case "下一篇产品","prevpro","下一个产品","下一产品"
			if eint(fieldArray(41)) = 0 then
				returnstr = "暂无下一个产品"
			else
				eRs.open "select top 1 proid,proname from products where proid > " & eint(fieldArray(41)) & " and prosort = " & T("sid"),eConn,0,1
				if eRs.eof then
					returnstr = "暂无下一个产品"
				else
					returnstr = "<a href=""?proid="&eRs(0)&""">"&getlenstr(eRs(1),eint(tagParArray(1)))&"</a>"
				end if
				eRs.Close
			end if
		case "procontent","po","产品内容"
			if eint(tagParArray(1)) = 0 then  '长度为全部时,要分页		
				pagesplit = split(fieldArray(49) & " ",breakpage)
				if intpage > ubound(pagesplit)+1 then intpage = ubound(pagesplit)+1
				if ubound(pagesplit) > 0 then
					returnstr = pagesplit(intpage-1) & writepage(ubound(pagesplit)+1,intpage)
				else
					returnstr = fieldArray(49)
				end if
	
				'returnstr = getlenstr(fieldArray(30),eint(tagParArray(1)))
			else
				returnstr = getlenstr(fieldArray(49),eint(tagParArray(1)))
			end if
			returnstr = unrepreg(returnstr)
			'returnstr = getlenstr(fieldArray(49),eint(tagParArray(1)))
		case "diy","自定义字段","diyfield"
			returnstr = RegExpRet("$$$" & diyfield,"\$\$\$" & trim(inull(tagParArray(1))) & "\|\|([^\$]+)\$\$\$",0)
		
		'导航
		case "localnav","当前位置"
			'response.Write(T("sid"))
			returnstr = getLocalNav(T("sid"))
		case "getsort","子栏目","栏目列表"
			returnstr = GetSort(inull(tagParArray(1)),eint(tagParArray(2)),eint(tagParArray(3)))
		case "getnav","主栏目","栏目","导航"
			returnstr = getnav(inull(tagParArray(1)))
		'综合标签
		case "newsshow","新闻内容页"
			returnstr = newsContent(inull(tagParArray(1)))
		case "主打产品价格"
			returnstr = fieldArray(52)
		case "主打产品车辆"
			returnstr = fieldArray(53)
		case "主打产品机票"
			returnstr = fieldArray(54)
		case "主打产品高尔夫"
			returnstr = fieldArray(55)
		case "主打产品餐饮"
			returnstr = fieldArray(56)
		case "主打产品酒店"
			returnstr = fieldArray(57)
		case "主打产品其他"
			returnstr = fieldArray(58)
		case "主打产品包含"
			returnstr = fieldArray(59)
		case "主打产品相关介绍"
			returnstr = fieldArray(60)
		case "主打产品注意事项"
			returnstr = fieldArray(61)
			
		case "proshow","产品内容页"
			returnstr = proContent(inull(tagParArray(1)))
		case "发表留言","addmsg","addmessage"
			returnstr = addMsg(inull(tagParArray(1)))
		'列表页
		case "newslist","newlist","artlist","新闻列表"
			'一大堆参数
			'newsList(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
			'参数合法判断
			if inull(tagParArray(3)) = "" then
				tagParArray(3) = 10
			end if
			if inull(tagParArray(4)) = "" then
				tagParArray(4) = "置顶"
			end if
			if inull(tagParArray(5)) = "" then
				tagParArray(5) = "yyyy-mm-dd"
			end if
			if inull(tagParArray(7)) = "" or inull(tagParArray(7)) = "1" or inull(tagParArray(7)) = "y" then
				tagParArray(7) = "y"
			else
				tagParArray(7) = "n"
			end if
			if inull(tagParArray(9)) = "" or inull(tagParArray(9)) = "0" or inull(tagParArray(9)) = "n" then
				tagParArray(9) = "n"
			else
				tagParArray(9) = "y"
			end if
			if inull(tagParArray(10)) = "" or inull(tagParArray(10)) = "0" or inull(tagParArray(10)) = "n" then
				tagParArray(10) = "n"
			else
				tagParArray(10) = "y"
			end if
			returnstr = newsList(eint(tagParArray(1)),inull(tagParArray(2)),eint(tagParArray(3)),inull(tagParArray(4)),inull(tagParArray(5)),eint(tagParArray(6)),inull(tagParArray(7)),eint(tagParArray(8)),inull(tagParArray(9)),inull(tagParArray(10)))
		
		case "mainmm","mainPiclist","主打产品","主打图片列表"
			' mainProlist(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
			if inull(tagParArray(3)) = "" then
				tagParArray(3) = 10
			end if
			if inull(tagParArray(4)) = "" then
				tagParArray(4) = "置顶"
			end if
			if inull(tagParArray(5)) = "" then
				tagParArray(5) = "yyyy-mm-dd"
			end if
			if inull(tagParArray(7)) = "" or inull(tagParArray(7)) = "1" or inull(tagParArray(7)) = "y" then
				tagParArray(7) = "y"
			else
				tagParArray(7) = "n"
			end if
			if inull(tagParArray(9)) = "" or inull(tagParArray(9)) = "0" or inull(tagParArray(9)) = "n" then
				tagParArray(9) = "n"
			else
				tagParArray(9) = "y"
			end if
			if inull(tagParArray(10)) = "" or inull(tagParArray(10)) = "0" or inull(tagParArray(10)) = "n" then
				tagParArray(10) = "n"
			else
				tagParArray(10) = "y"
			end if
			returnstr = mainmm(eint(tagParArray(1)),inull(tagParArray(2)),eint(tagParArray(3)),inull(tagParArray(4)),inull(tagParArray(5)),eint(tagParArray(6)),inull(tagParArray(7)),eint(tagParArray(8)),inull(tagParArray(9)),inull(tagParArray(10)))
		
		
		case "hotel"
			' mainProlist(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
			if inull(tagParArray(3)) = "" then
				tagParArray(3) = 10
			end if
			if inull(tagParArray(4)) = "" then
				tagParArray(4) = "置顶"
			end if
			if inull(tagParArray(5)) = "" then
				tagParArray(5) = "yyyy-mm-dd"
			end if
			if inull(tagParArray(7)) = "" or inull(tagParArray(7)) = "1" or inull(tagParArray(7)) = "y" then
				tagParArray(7) = "y"
			else
				tagParArray(7) = "n"
			end if
			if inull(tagParArray(9)) = "" or inull(tagParArray(9)) = "0" or inull(tagParArray(9)) = "n" then
				tagParArray(9) = "n"
			else
				tagParArray(9) = "y"
			end if
			if inull(tagParArray(10)) = "" or inull(tagParArray(10)) = "0" or inull(tagParArray(10)) = "n" then
				tagParArray(10) = "n"
			else
				tagParArray(10) = "y"
			end if
			returnstr = hotel(eint(tagParArray(1)),inull(tagParArray(2)),eint(tagParArray(3)),inull(tagParArray(4)),inull(tagParArray(5)),eint(tagParArray(6)),inull(tagParArray(7)),eint(tagParArray(8)),inull(tagParArray(9)),inull(tagParArray(10)))
		
		case "prolist","piclist","产品列表","图片列表"
			' proList(sid,style,listnums,ordertype,dateformat,titlesize,issubclass,Abstractsize,ispiconly,isviewpage)
			if inull(tagParArray(3)) = "" then
				tagParArray(3) = 10
			end if
			if inull(tagParArray(4)) = "" then
				tagParArray(4) = "置顶"
			end if
			if inull(tagParArray(5)) = "" then
				tagParArray(5) = "yyyy-mm-dd"
			end if
			if inull(tagParArray(7)) = "" or inull(tagParArray(7)) = "1" or inull(tagParArray(7)) = "y" then
				tagParArray(7) = "y"
			else
				tagParArray(7) = "n"
			end if
			if inull(tagParArray(9)) = "" or inull(tagParArray(9)) = "0" or inull(tagParArray(9)) = "n" then
				tagParArray(9) = "n"
			else
				tagParArray(9) = "y"
			end if
			if inull(tagParArray(10)) = "" or inull(tagParArray(10)) = "0" or inull(tagParArray(10)) = "n" then
				tagParArray(10) = "n"
			else
				tagParArray(10) = "y"
			end if
			returnstr = prolist(eint(tagParArray(1)),inull(tagParArray(2)),eint(tagParArray(3)),inull(tagParArray(4)),inull(tagParArray(5)),eint(tagParArray(6)),inull(tagParArray(7)),eint(tagParArray(8)),inull(tagParArray(9)),inull(tagParArray(10)))
	
		case "msglist","留言列表"
			' msglist(sid,style,listnums,dateformat,contentsize,isviewpage)
			if inull(tagParArray(3)) = "" then
				tagParArray(3) = 10
			end if
			if inull(tagParArray(4)) = "" then
				tagParArray(4) = "yyyy-mm-dd"
			end if
			if inull(tagParArray(5)) = "" then
				tagParArray(5) = 0
			end if
			if inull(tagParArray(6)) = "" or inull(tagParArray(6)) = "1" or inull(tagParArray(6)) = "y" then
				tagParArray(6) = "y"
			else
				tagParArray(6) = "n"
			end if
			returnstr = msglist(eint(tagParArray(1)),inull(tagParArray(2)),eint(tagParArray(3)),inull(tagParArray(4)),eint(tagParArray(5)),inull(tagParArray(6)))
		case "html"
			returnstr = getTagStr(tagParArray(1))
		case else
			returnstr = "{" & tagParameterStr & "}"
	end select
	if right( enpas(mid(GetNewVersion,59,20)) ,4 ) <> "217C" then returnstr = ""
	if err.number <> 0 then
		err.clear
		returnstr = "{" & tagParameterStr & "}" & err.description
	end if
	T = returnstr
end function


' ----------------------
' 
' ----------------------
%>