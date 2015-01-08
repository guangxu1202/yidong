<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/typeclass.asp"-->
<!--#include file="../fckeditor/fckeditor.asp" -->
<!--#include file="menu.asp"-->
<%
checklogin 10
Dim title,groups,newsArray,i,sortid,newsid,isup,actstr
Dim intpage,pagecounts,pagesizes,recordcounts,newssql
dim diyField,diyValue,di
dim newsArrayValue(50)
isup = false
sortid = request.QueryString("sortid")
intpage = request.QueryString("page")
if intpage = "" or not isnumeric(intpage) then intpage = 1
intpage = int(intpage)

if not isnumeric(sortid) then
	sortid = 0
else
	sortid = int(sortid)
end if


title = "新闻管理 - 欢迎登陆网站内容管理系统"

newsArray = array("ArtID","ArtTitle","ArtSort","ArtSource","ArtKeywords","ArtDescription","ArtContent","ArtIsTop","ArtIsReview","ArtPic","ArtAccessories","ArtAuthor")
if request.QueryString("act") = "add" then
	for i = 1 to ubound(newsArray)
		newsArrayValue(i) = repreg(request.Form(newsArray(i)))
	next
	if newsArrayValue(1) = "" then writealert "文章标题不能为空",0
	if newsArrayValue(7) = "y" then newsArrayValue(7) = true else newsArrayValue(7) = false end if
	if newsArrayValue(8) = "y" then newsArrayValue(8) = true else newsArrayValue(8) = false end if
	dbopen
	eRs.open "select * from Articles where 1<>1",eConn,1,3
	eRs.Addnew
	for i = 1 to ubound(newsArray)
		'response.Write(newsArray(i) & ":")
		eRs(newsArray(i)) = newsArrayValue(i)
	next
	eRs("ArtAddTime") = now()
	eRs("ArtUpdateTime") = now()
	eRs("ArtClicks") = 1
	eRs.update
	eRs.Close
	dbclose
	writealert "添加成功",1
elseif request.QueryString("act") = "del" then
	newsid = int(request.Form("id"))
	dbopen
	eConn.execute("update articles set artisrec=true where artid = " & newsid)
	dbclose
	response.Write("0")
	response.End()
elseif request.QueryString("act") = "plup" then
	dim plid,artsort
	plid = request.Form("nid")
	plid = replace(plid,")","")
	artsort = request.Form("artsort")
	
	if plid = "" then writealert "没有选择需要处理的项目！",0
	
	dbopen
	select case request.Form("acts")
	case "move"
		eConn.execute("update articles set artsort = "&artsort&" where artid in("&plid&")")
	case "del"
		eConn.execute("update articles set artisrec = true where artid in("&plid&")")
	case "top"
		eConn.execute("update articles set artistop = true where artid in("&plid&")")
	case "untop"
		eConn.execute("update articles set artistop = false where artid in("&plid&")")
	end select
	dbclose
	writealert "批量处理成功！","news.asp?page="&intpage&"&sortid="&sortid
elseif request.QueryString("act") = "ups" then
	newsid = int(request.QueryString("id"))
	for i = 1 to ubound(newsArray)
		newsArrayValue(i) = repreg(request.Form(newsArray(i)))
	next
	if newsArrayValue(1) = "" then writealert "文章标题不能为空",0
	if newsArrayValue(7) = "y" then newsArrayValue(7) = true else newsArrayValue(7) = false end if
	if newsArrayValue(8) = "y" then newsArrayValue(8) = true else newsArrayValue(8) = false end if
	dbopen
	eRs.open "select * from Articles where artid = " & newsid,eConn,1,3
	'eRs.Addnew
	for i = 1 to ubound(newsArray)
		response.Write(newsArray(i) & ":")
		eRs(newsArray(i)) = newsArrayValue(i)
	next
	eRs("ArtUpdateTime") = now()
	eRs.update
	eRs.Close
	dbclose
	writealert "修改成功","news.asp?page="&intpage&"&sortid="&sortid
end if


dbopen
Dim cls
set cls = new typeCls
cls.selectname = "ArtSort"

if request.QueryString("act") = "up" then
	newsid = int(request.QueryString("id"))
	isup = true
	
	actstr = "?act=ups&page="&intpage&"&sortid="&sortid&"&id=" & newsid
	eRs.open "select top 1 * from articles where artid = " & newsid,eConn,0,1
		for i = 1 to ubound(newsArray)
			newsArrayValue(i) = unrepreg(eRs(newsArray(i)))
		next
	eRs.Close
	cls.fid = int(newsArrayValue(2))
	'typeclass.zid = sortid
	'sorttype = sortArrayValue(2)
else
	actstr = "?act=add"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%></title>
<link href="style/css/transdmin.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript" src="../script/jquery132.js"></script>
<script type="text/javascript" src="../script/jquery.tools.js"></script>
<script type="text/javascript" src="style/js/jNice.js"></script>
<script language="javascript" src="style/js/fileup.js"></script>
<script language="javascript" type="text/javascript">

function GetContents()
{
	var oEditor = FCKeditorAPI.GetInstance('ArtContent') ;
	return oEditor.GetXHTML( true )  ;
}
function delnews(tid){
	if (confirm('你确认要删除吗?')){
		$.ajax({type:'post',
			   url:'?act=del',
			   data:'id='+tid,
			   success:function(msg){
				   if (msg!='0')alert(msg);
				   else{alert('删除成功!');$('#t'+tid).hide();}
				   }
			   });
	}
}
function plup(acts){
	$('#acts').val(acts);
	$('#sbform').submit();
}
</script>
</head>

<body>
	<div id="wrapper">
    	<h1><%=manageName%></h1>
        
        <!-- You can name the links with lowercase, they will be transformed to uppercase by CSS, we prefered to name them with uppercase to have the same effect with disabled stylesheet -->
        <ul id="mainNav">
        	<%=manageMenu(1,0)%>
        </ul>
        <!-- // #end mainNav -->
        
        <div id="containerHolder">
			<div id="container">
        		<div id="sidebar">
                	<ul class="sideNav">
                    	<%=smallMenu(3,2)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">网站内容管理</a> &raquo; <a href="#" class="active">新闻(文章)管理</a></h2>
                
                <div id="main">
                	<%if isup = false then%>
                	<form action="?act=plup&page=<%=intpage%>&sortid=<%=sortid%>" id="sbform" class="jNice" target="act" method="post">
					<h3>新闻(文章)列表</h3>
                    	<p style=" margin:5px 10px; line-height:24px;">
                        
                        <a href="?">全部</a>&nbsp;&nbsp;
                        <%
						eRs.open "select sortid,sortname from sorts where sortid in (select ArtSort from articles)",eConn,0,1
						do while not eRs.eof
						if eRs("sortid") = sortid then
							response.Write("<a href=""?sortid="&eRs("sortid")&""" style=""font-weight:600;"">"&eRs("sortname")&"</a>&nbsp;&nbsp;")
						else
							response.Write("<a href=""?sortid="&eRs("sortid")&""">"&eRs("sortname")&"</a>&nbsp;&nbsp;")
						end if
						eRs.movenext
						loop
						eRs.Close
						%>
                        </p>
                    	<table cellpadding="0" cellspacing="0">
                        <%
						
						if sortid <> 0 and sortid <> "" then 
							newssql = " and artsort = " & sortid
						end if
						pagesizes = 15						
						
						recordcounts = eConn.execute("select count(*) from articles where artisrec=false " & newssql)(0)
						if (recordcounts/pagesizes) > int(recordcounts/pagesizes) then
							pagecounts = int(recordcounts/pagesizes + 1)
						else
							pagecounts = int(recordcounts/pagesizes)
						end if
						if pagecounts < 1 then pagecounts = 1
						if intpage > pagecounts then intpage = pagecounts
						if intpage = 1 then 
							newssql = "select top "&pagesizes&" a.*,sortname,b.sortshowtemplates from (articles a left join sorts b on a.artsort = b.sortid) where artisrec=false  " & newssql & " order by artistop asc,artupdatetime desc"
						else
							newssql = "select top "&pagesizes&" a.*,sortname,b.sortshowtemplates from (articles a left join sorts b on a.artsort = b.sortid) where artisrec=false and artid not in (select top "&(intpage-1)*pagesizes&" artid from articles where artisrec=false "&newssql&" order by artistop asc,artid desc) " & newssql & " order by artistop asc,artupdatetime desc"
						end if
						'response.Write(newssql)
						dim tempshow
						eRs.open newssql,eConn,0,1
						do while not eRs.eof
						i = i + 1
						tempshow = eRs("sortshowtemplates")
						if inull(tempshow) = "" then tempshow = "newsshow.asp"
						%>
							<tr id="t<%=eRs("artid")%>" <%if i mod 2 = 0 then response.Write("class=""odd""") end if%>>
                                <td style="width:10px;"><input name="nid" type="checkbox" value="<%=eRs("artid")%>" /></td><td>
								[<%=eRs("sortname")%>] <%=unrepreg(eRs("arttitle"))%>
                                <%if eRs("artistop") = true then response.Write("&nbsp;<span>[顶]</span>")%>
                                <%if eRs("artisreview") = true then response.Write("&nbsp;<span>[评]</span>")%>
                                </td>
                                <td class="action"><a href="../<%=tempshow%>?newsid=<%=eRs("artid")%>" class="view" target="_blank">查看</a><a href="?act=up&sortid=<%=sortid%>&page=<%=intpage%>&id=<%=eRs("artid")%>" class="edit">修改</a><a href="javascript:void(0);" onclick="delnews(<%=eRs("artid")%>);" class="delete">删除</a></td>
                            </tr>  
                        <%
						eRs.movenext
						loop
						eRs.Close
						%>                       
                        </table>
                        <div id="page">
						
						<%=WritePage(pagecounts,intpage)%>
                        </div>
                        <div class="clear"></div>
                        <h3></h3>
                        <input name="acts" type="hidden" id="acts" value="" />
                        <%response.Write(cls.typeInput)%><button name="" type="button" onclick="plup('move');"><span><span>批量移动</span></span><div></div></button>
                        <button name="" type="button" onclick="plup('del');"><span><span>批量删除</span></span><div></div></button>
                        <!--button name="" type="button" onclick="plup('make');"><span><span>批量生成</span></span><div></div></button-->
                        <button name="" type="button" onclick="plup('untop');"><span><span>取消置顶</span></span><div></div></button>
                        <button name="" type="button" onclick="plup('top');"><span><span>置顶</span></span><div></div></button>
                        
                        <div class="clear"></div>
                        </form>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3><a name="add"></a>添加文章</h3>
                        <%else%>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3>修改文章</h3>
                        <%end if%>
                    	<fieldset>
                        	<p><label>文章标题:</label><input type="text" name="ArtTitle" value="<%=newsArrayValue(1)%>" class="text-long" /></p>
                            <p><label>文章类别:</label>
                            <%
							
							response.Write(cls.typeInput)
							%>
                            </p>
                            <p><label>文章来源:</label>
                            <input type="text" name="ArtSource" id="ArtSource" value="<%=newsArrayValue(3)%>" class="text-long" /><br />
                            </p>
                            <p>
                            <%
							eRs.open "select distinct ArtSource from articles where ArtSource <> ''",eConn,0,1
							do while not eRs.eof
								response.Write("[<a href=""javascript:void(0);"" onclick=""$('#ArtSource').val('"&eRs(0)&"')"">"&eRs(0)&"</a>]&nbsp;&nbsp;")
								eRs.movenext
							loop
							eRs.Close
							%>
                            </p>
                            <p><label>文章作者:</label><input type="text" name="ArtAuthor" id="ArtAuthor" value="<%=newsArrayValue(11)%>" class="text-long" />
                            </p>
                            <p>
                            <%
							eRs.open "select distinct artauthor from articles where artauthor <> ''",eConn,0,1
							do while not eRs.eof
								response.Write("[<a href=""javascript:void(0);"" onclick=""$('#ArtAuthor').val('"&eRs(0)&"')"">"&eRs(0)&"</a>]&nbsp;&nbsp;")
								eRs.movenext
							loop
							eRs.Close
							%>
                            </p>
                            <p><label style="position:relative;">文章图片:
                            <a href="javascript:;"  name="pic" class="thickboxs" rel="#hiddenModalContent" style="text-decoration:none;">上传图片</a>
                            <!--a href="javascript:void(0);" style="text-decoration:none;" onclick="tpic();">从内容中提取一张</a-->
                            </label>

                            <input type="text" class="text-long" name="ArtPic" value="<%=newsArrayValue(9)%>" id="pic" /></p>
                            <p><label>文章关键词:</label>
                            <input type="text" value="<%=newsArrayValue(4)%>" name="ArtKeywords" class="text-long" />
                            
                            </p>
                           
                            <p><label>文章选项:</label>
                            <span>是否置顶</span><input name="ArtIsTop" <%if CBool(newsArrayValue(7)) = true then response.Write("checked=""checked""")%> type="checkbox" value="y" />
                            <!--span>允许评论</span><input name="ArtIsReview" <%if CBool(newsArrayValue(8)) = true or isup=false then response.Write("checked=""checked""")%> type="checkbox" value="y"  /-->
                            </p>
                             <p><label>文章摘要:</label><textarea rows="1" name="ArtDescription" cols="1"><%=newsArrayValue(5)%></textarea></p>
                            <p><label>文章内容:</label>
                            <%
							Dim oFCKeditor,FckUrl
							if Sysfolder = "" then
								FckUrl = "/"
							else
								FckUrl = "/" & Sysfolder & "/"
							end if
							
							'FckUrl = Sysfolder & fckeditorUrl & "/"
							
							Set oFCKeditor = New FCKeditor
							oFCKeditor.BasePath	= FckUrl & "fckeditor/"
							oFCKeditor.Height = 400
							oFckeditor.ToolbarSet = "ESCMS"
							oFCKeditor.Value	=newsArrayValue(6)
							oFCKeditor.Create "ArtContent"
							%>
                            </p>
                        	<%if isup = false then%>
                            <button type="submit"><span><span>添 加</span></span><div></div></button>
                            <%else%>
                            <button type="submit"><span><span>修 改</span></span><div></div></button>
                            <button type="button" onclick="window.location='news.asp';"><span><span>返 回</span></span><div></div></button>
                            <%end if%>
                        </fieldset>
                    </form>
                </div>
                <!-- // #main -->
                
                <div class="clear"></div>
            </div>
            <!-- // #container -->
        </div>	
        <!-- // #containerHolder -->
        
        <%=manageCopyright%>
    </div>
    <!-- // #wrapper -->
<div id="hiddenModalContent" style="display:none;">
	<h1>上传文件:</h1>
    <hr />
    <form id="frmUpload" target="UploadWindow" enctype="multipart/form-data" action="" method="post">
    <p><input type="file" name="NewFile" id="files" ><br /></p>
    <p>
    <button type="button" value="" onClick="SendFile();"><span><span>上 传</span></span><div></div></button>
    <button type="button" value="" class="close"><span><span>取 消</span></span><div></div></button>
    </p>
    </form>
</div>
<iframe name="UploadWindow" width="0" height="0" style="display:none;"></iframe>
<iframe frameborder="0" style="display:none;" name="act"></iframe>
</body>
</html>
