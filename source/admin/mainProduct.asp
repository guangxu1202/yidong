<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/typeclass.asp"-->
<!--#include file="../fckeditor/fckeditor.asp" -->
<!--#include file="menu.asp"-->
<%
checklogin 10
Dim title,groups,ProArray,i,sortid,newsid,isup,actstr
Dim intpage,pagecounts,pagesizes,recordcounts,newssql
dim diyField,diyValue,di
dim ProArrayValue(50)
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


title = "产品管理 - 欢迎登陆网站内容管理系统"
'SortDiyField
ProArray = array("ProID","ProName","ProSort","ProsmallPic","ProBigPic","ProIsTop","ProIsReview","ProKeywords","Prodescription","ProContent","proPrice","proCar","proTicket","proGolf","proEat","proHotel","proElse","proIn","prointro","proNote")
if request.QueryString("act") = "add" then
	for i = 1 to ubound(ProArray)
		ProArrayValue(i) = repreg(request.Form(ProArray(i)))
	next
	if ProArrayValue(1) = "" then writealert "产品标题不能为空",0
	if ProArrayValue(5) = "y" then ProArrayValue(5) = true else ProArrayValue(5) = false end if
	if ProArrayValue(6) = "y" then ProArrayValue(6) = true else ProArrayValue(6) = false end if
	
	dbopen
	if ProArrayValue(2) <> "" then	
	 	diyField = eConn.execute("select top 1 SortDiyField from sorts where sortid = " & ProArrayValue(2))(0)
		diyField = split(inull(diyField),"|")
		for di = 0 to ubound(diyField)
			diyValue = diyValue & split(diyField(di),":")(0) & "||" & repreg(request.Form("diy_" & split(diyField(di),":")(0))) & "$$$"
		next
	end if

	
	eRs.open "select * from mainProduct where 1<>1",eConn,1,3
	eRs.Addnew
	for i = 1 to ubound(ProArray)
		'response.Write(ProArray(i) & ":")
		eRs(ProArray(i)) = ProArrayValue(i)
	next
	eRs("proAddTime") = now()
	eRs("ProUpdateTime") = now()
	eRs("proClicks") = 1
	eRs("proDiyValue") = diyValue
	eRs.update
	eRs.Close
	dbclose
	writealert "添加成功",1
elseif request.QueryString("act") = "plup" then
	dim plid,artsort
	plid = request.Form("nid")
	plid = replace(plid,")","")
	artsort = request.Form("prosort")
	
	if plid = "" then writealert "没有选择需要处理的项目！",0
	
	dbopen
	select case request.Form("acts")
	case "move"
		eConn.execute("update mainProduct set prosort = "&artsort&" where proid in("&plid&")")
	case "del"
		eConn.execute("update mainProduct set proisrec = true where proid in("&plid&")")
	case "top"
		response.Write("update mainProduct set proistop = true where proid in("&plid&")")
		eConn.execute("update mainProduct set proistop = true where proid in("&plid&")")
	case "untop"
		eConn.execute("update mainProduct set proistop = false where proid in("&plid&")")
	end select
	dbclose
	writealert "批量处理成功！","mainProduct.asp?page="&intpage&"&sortid="&sortid
elseif request.QueryString("act") = "del" then
	newsid = int(request.Form("id"))
	dbopen
	eConn.execute("update mainProduct set proisrec=true where proid = " & newsid)
	dbclose
	response.Write("0")
	response.End()
elseif request.QueryString("act") = "ups" then
	newsid = int(request.QueryString("id"))
	for i = 1 to ubound(ProArray)
		ProArrayValue(i) = repreg(request.Form(ProArray(i)))
	next
	if ProArrayValue(1) = "" then writealert "文章标题不能为空",0
	if ProArrayValue(5) = "y" then ProArrayValue(5) = true else ProArrayValue(5) = false end if
	if ProArrayValue(6) = "y" then ProArrayValue(6) = true else ProArrayValue(6) = false end if
	dbopen
	if ProArrayValue(2) <> "" then	
	 	diyField = eConn.execute("select top 1 SortDiyField from sorts where sortid = " & ProArrayValue(2))(0)
		diyField = split(inull(diyField),"|")
		for di = 0 to ubound(diyField)
			diyValue = diyValue & split(diyField(di),":")(0) & "||" & repreg(request.Form("diy_" & split(diyField(di),":")(0))) & "$$$"
		next
	end if
	'response.Write(diyValue)
	'response.End()
	
	eRs.open "select * from mainProduct where proid = " & newsid,eConn,1,3
	'eRs.Addnew
	for i = 1 to ubound(ProArray)
		'response.Write(ProArray(i) & ":")
		eRs(ProArray(i)) = ProArrayValue(i)
	next
	eRs("ProUpdateTime") = now()
	eRs("proDiyValue") = diyValue
	eRs.update
	eRs.Close
	dbclose
	writealert "修改成功","mainProduct.asp?page="&intpage&"&sortid="&sortid
end if


dbopen
Dim cls
set cls = new typeCls
cls.selectname = "ProSort"


if request.QueryString("act") = "up" then
	newsid = int(request.QueryString("id"))
	isup = true
	
	actstr = "?act=ups&page="&intpage&"&sortid="&sortid&"&id=" & newsid
	eRs.open "select top 1 * from mainProduct where proid = " & newsid,eConn,0,1
		for i = 1 to ubound(ProArray)
			ProArrayValue(i) = unrepreg(eRs(ProArray(i)))
		next
	eRs.Close
	cls.fid = int(ProArrayValue(2))
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
function cgtype(){
	//alert($("#sorttype").val());
	//return;
	$('#diyfield').load('diyfield.asp?sid='+$("#sorttype").val()+'&nid=<%=newsid%>&t='+Math.random());
}
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
                    	<%=smallMenu(3,5)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">网站内容管理</a> &raquo; <a href="#" class="active">主打产品管理</a></h2>
                
                <div id="main">
                	<%if isup = false then%>
                	<form action="?act=plup&page=<%=intpage%>&sortid=<%=sortid%>" class="jNice" id="sbform" target="act" method="post">
					<h3>主打产品列表</h3>
                    	<p style=" margin:5px 10px; line-height:24px;">
                        
                        <a href="?">全部</a>&nbsp;&nbsp;
                        <%
						eRs.open "select sortid,sortname from sorts where sortid in (select proSort from mainProduct)",eConn,0,1
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
							newssql = " and prosort = " & sortid
						end if
						pagesizes = 15						
						
						recordcounts = eConn.execute("select count(*) from mainProduct where proisrec=false " & newssql)(0)
						if (recordcounts/pagesizes) > int(recordcounts/pagesizes) then
							pagecounts = int(recordcounts/pagesizes + 1)
						else
							pagecounts = int(recordcounts/pagesizes)
						end if
						if pagecounts < 1 then pagecounts = 1
						if intpage > pagecounts then intpage = pagecounts
						if intpage = 1 then 
							newssql = "select top "&pagesizes&" a.*,sortname,b.sortshowtemplates from (mainProduct a left join sorts b on a.prosort = b.sortid) where  proisrec=false  " & newssql & " order by proistop asc,proid desc"
						else
							newssql = "select top "&pagesizes&" a.*,sortname,b.sortshowtemplates from (mainProduct a left join sorts b on a.prosort = b.sortid) where  proisrec=false and proid not in (select top "&(intpage-1)*pagesizes&" proid from mainProduct where  proisrec=false "&newssql&" order by proistop asc,proid desc) " & newssql & " order by proistop asc,proid desc"
						end if
						'response.Write(newssql)
						dim tempshow
						eRs.open newssql,eConn,0,1
						do while not eRs.eof
						i = i + 1
						tempshow = eRs("sortshowtemplates")
						if inull(tempshow) = "" then tempshow = "newsshow.asp"
						%>
							<tr id="t<%=eRs("proid")%>" <%if i mod 2 = 0 then response.Write("class=""odd""") end if%>>
                                <td style="width:10px;"><input name="nid" type="checkbox" value="<%=eRs("proid")%>" /></td><td>
								[<%=eRs("sortname")%>] <%=unrepreg(eRs("proname"))%>
                                <%if eRs("proistop") = true then response.Write("&nbsp;<span>[顶]</span>")%>
                                <%if eRs("proisreview") = true then response.Write("&nbsp;<span>[评]</span>")%>
                                </td>
                                <td class="action"><a href="../<%=tempshow%>?mproid=<%=eRs("proid")%>" class="view" target="_blank">查看</a><a href="?act=up&sortid=<%=sortid%>&page=<%=intpage%>&id=<%=eRs("proid")%>" class="edit">修改</a><a href="javascript:void(0);" onclick="delnews(<%=eRs("proid")%>);" class="delete">删除</a></td>
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
                        <!--
                        <input name="acts" type="hidden" id="acts" value="" />
                        <%response.Write(cls.typeInput)%><button name="" type="button" onclick="plup('move');"><span><span>批量移动</span></span><div></div></button>
						-->
                        <button name="" type="button" onclick="plup('del');"><span><span>批量删除</span></span><div></div></button>
                        <!--button name="" type="button" onclick="plup('make');"><span><span>批量生成</span></span><div></div></butto
                        <button name="" type="button" onclick="plup('untop');"><span><span>取消置顶</span></span><div></div></button>
                        <button name="" type="button" onclick="plup('top');"><span><span>置顶</span></span><div></div></button>n-->
                        
                        
                        <div class="clear"></div>
                        </form>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3><a name="add"></a>添加产品</h3>
                        <%else%>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3>修改产品</h3>
                        <%end if%>
                    	<fieldset>
                        	<p><label>产品名称:</label><input type="text" name="ProName" value="<%=ProArrayValue(1)%>" class="text-long" /></p>
                            <p><label>产品类别:</label>
                            <%
							cls.click = " id=""sorttype"" onchange=""cgtype();"" "
							response.Write(cls.typeInput)
							%>
                            </p>

                            <p><label style="position:relative;">产品图片:
                            <a href="javascript:;"  name="spic" class="thickboxs" rel="#hiddenModalContent" style="text-decoration:none;">上传图片</a>
                            </label>
							<input type="text" class="text-long" name="ProsmallPic" value="<%=ProArrayValue(3)%>" id="spic" /></p>
							
							<p><label>价格:</label><input type="text" name="proPrice" value="<%=ProArrayValue(10)%>" class="text-long" />元</p>
							
							<p><label>车辆:</label><input type="text" name="proCar" value="<%=ProArrayValue(11)%>" class="text-long" /></p>
							
							<p><label>机票:</label><input type="text" name="proTicket" value="<%=ProArrayValue(12)%>" class="text-long" /></p>
							
							<p><label>高尔夫:</label><input type="text" name="proGolf" value="<%=ProArrayValue(13)%>" class="text-long" /></p>
							
							<p><label>餐饮:</label><input type="text" name="proEat" value="<%=ProArrayValue(14)%>" class="text-long" /></p>
							
							<p><label>酒店:</label><input type="text" name="proHotel" value="<%=ProArrayValue(15)%>" class="text-long" /></p>
							
							<p><label>其他:</label><textarea rows="1" name="proElse" cols="1"><%=ProArrayValue(16)%></textarea></p>
							
							<p><label>包含/不包含:</label><textarea rows="1" name="proIn" cols="1"><%=ProArrayValue(17)%></textarea></p>
							
							<p><label>相关介绍:</label><textarea rows="1" name="prointro" cols="1"><%=ProArrayValue(18)%></textarea></p>
							
							<p><label>注意事项:</label><textarea rows="1" name="proNote" cols="1"><%=ProArrayValue(19)%></textarea></p>

							<!--
                            <input type="text" class="text-long" name="ProsmallPic" value="<%=ProArrayValue(3)%>" id="spic" /></p>
                             <p><label style="position:relative;">产品大图:
                            <a href="javascript:;"  name="bpic" class="thickboxs" rel="#hiddenModalContent" style="text-decoration:none;">上传图片</a>
                            </label>
							
                            <input type="text" class="text-long" name="ProBigPic" value="<%=ProArrayValue(4)%>" id="bpic" /></p>
                            <p><label>产品关键词:</label>
                            <input type="text" value="<%=ProArrayValue(7)%>" name="ProKeywords" class="text-long" />
                            </p>
                           -->
                            <p><label>产品选项:</label>
                            <span>是否置顶</span><input name="ProIsTop" <%if CBool(ProArrayValue(5)) = true then response.Write("checked=""checked""")%>  type="checkbox" value="y" />
							
                            <!--span>允许评论</span><input name="ProIsReview" <%if CBool(ProArrayValue(6)) = true or isup=false then response.Write("checked=""checked""")%>  type="checkbox" value="y" checked="checked" /-->
                            </p>
                            <div id="diyfield">
                            </div>
							<!--
                             <p><label>产品简介:</label><textarea rows="1" name="Prodescription" cols="1"><%=ProArrayValue(8)%></textarea></p>
							 -->
                            <p><label>行程:</label>
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
							oFCKeditor.Value	=ProArrayValue(9)
							oFCKeditor.Create "ProContent"
							%>
                            </p>
                        	<%if isup = false then%>
                            <button type="submit"><span><span>添 加</span></span><div></div></button>
                            <%else%>
                            <button type="submit"><span><span>修 改</span></span><div></div></button>
                            <button type="button" onclick="window.location='mainProduct.asp';"><span><span>返 回</span></span><div></div></button>
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
<%
if isup = true then
response.Write("<script>cgtype();</script>")
end if
%>
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
