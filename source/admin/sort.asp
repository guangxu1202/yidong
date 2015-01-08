<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/typeclass.asp"-->
<!--#include file="../fckeditor/fckeditor.asp" -->
<!--#include file="menu.asp"-->
<%
checklogin 5
Dim title,groups
dim sortArray,i,isup,actstr,sortid,sorttype
dim sortArrayValue(50)

title = "栏目管理 - 网站内容管理系统"
isup = false
sorttype = 0

'已定义数组变量请勿改变顺序。。偷懒这么写的，呃
sortArray = array("sortID","sortname","sorttype","sortfid","SortIsNav","SortIsList","sortorder","sortkeywords","sortdescription","sortpic","SortListTemplates","SortShowTemplates","sortcontent","SortDiyField")
'sortArrayValue = sortArray  '完成值数组定义
	
if request.QueryString("act") = "add" then

	for i = 1 to ubound(sortArray)
		sortArrayValue(i) = request.Form(sortArray(i))
	next
	if request.Form(sortArray(4)) = "y" then sortArrayValue(4) = true else sortArrayValue(4) = false end if
	if request.Form(sortArray(5)) = "y" then sortArrayValue(5) = true else sortArrayValue(5) = false end if
	if sortArrayValue(1) = "" then writealert "栏目名称不能为空",0
	'response.Write("select top 1 * from sorts where sortname = '"&sortArrayValue(1)&"' and SortFID = " & sortArrayValue(3) & "<br />")
	dbopen
		eRs.open "select top 1 * from sorts where sortname = '"&sortArrayValue(1)&"' and SortFID = " & sortArrayValue(3),eConn,1,3
		if not eRs.eof then 
			eRs.Close
			dbclose
			writealert "同级栏目名称不能重复！",0
		end if
		eRs.addnew
			for i=1 to ubound(sortArray)
				'response.Write(sortArray(i) & ":" & sortArrayValue(i))
				eRs(sortArray(i)) = sortArrayValue(i)
			next
			eRs("sortaddtime") = now()
		eRs.update
		
		if sortArrayValue(2) = 1 then
			eConn.execute("insert into Alones (AloneName,AloneSort,AlonePic,AloneIsSort,AloneContent,AloneAddtime,AloneShowPage) values ('"&replace(sortArrayValue(1),"'","''")&"',"&eRs("sortid")&",'"&replace(sortArrayValue(9),"'","''")&"',"&eRs("sortid")&",'"&replace(sortArrayValue(12),"'","''")&"','"&now()&"','"&replace(sortArrayValue(11),"'","''")&"')")
		end if
		eRs.Close
	dbclose
	writealert "添加成功！",1
elseif request.QueryString("act") = "del" then
	sortid = request.form("id")
	if sortid = "" or not isnumeric(sortid) then 
		response.Write("ID不为数字或为空!")
		response.End()
	end if
	'on error resume next
	dbopen

	eRs.open "select top 1 1 from sorts where  sortfid = " & sortid,econn,0,1
	if not eRs.eof then
		eRs.Close
		dbclose
		response.Write("该目录下有子目录,不能删除.")
		response.End()
	else
		eRs.Close
	end if
	
	'删除栏目关联单页及其下属  后果太严重,暂不删除
	'eConn.execute("delete from alones where aloneissort = "&sortid&" or alonesort =  "&sortid)
	'eConn.execute("delete from articles where artsort =  "&sortid)
	'eConn.execute("delete from products where prosort =  "&sortid)
	
	eConn.execute("delete from sorts where sortid = " & sortid)
	response.Write("0")
	response.End()
	dbClose
elseif request.QueryString("act") = "ups" then
	
	sortid = request.QueryString("id")
	if sortid = "" or not isnumeric(sortid) then 
		writealert "ID不为数字或为空!",0
	end if
	for i = 1 to ubound(sortArray)
		sortArrayValue(i) = request.Form(sortArray(i))
	next
	if request.Form(sortArray(4)) = "y" then sortArrayValue(4) = true else sortArrayValue(4) = false end if
	if request.Form(sortArray(5)) = "y" then sortArrayValue(5) = true else sortArrayValue(5) = false end if
	if sortArrayValue(1) = "" then writealert "栏目名称不能为空",0
	'response.Write("select top 1 * from sorts where sortid = " & sortid & "<br />")
	dbopen
		
		'限制移动的目录同级不能重名
		eRs.open "select top 1 1 from sorts where sortid <> "&sortid&" and sortfid = "&sortArrayValue(3)&" and sortname ='"&sortArrayValue(1)&"'",eConn,1,3
			if not eRs.eof then
			eRs.Close
			dbclose
			writealert "同级栏目名称重复！",0
			end if
		eRs.Close		
		
		eRs.open "select top 1 * from sorts where sortid = " & sortid,eConn,1,3
		if eRs.eof then 
			eRs.Close
			dbclose
			writealert "未找到该栏目！",0
		end if
		'eRs.addnew
			for i=1 to ubound(sortArray)
				'response.Write(sortArray(i) & ":" & sortArrayValue(i))
				eRs(sortArray(i)) = sortArrayValue(i)
			next
			eRs("sortaddtime") = now()
		eRs.update
		eRs.Close
		if sortArrayValue(2) = 1 then
			'eConn.execute("update Alones set AloneSort,AloneIsSort,AloneShowPage where AloneIsSort = " & sortid)
		end if
	dbclose
	writealert "修改成功！","sort.asp"
end if
dbopen
Dim typeclass
set typeclass = new typeCls
if request("act")="up" then
	sortid = int(request.QueryString("id"))
	isup = true
	
	actstr = "?act=ups&id=" & sortid
	eRs.open "select top 1 * from sorts where sortid = " & sortid,eConn,0,1
		for i = 1 to ubound(sortArray)
			sortArrayValue(i) = eRs(sortArray(i))
		next
	eRs.Close
	typeclass.fid = sortArrayValue(3)
	typeclass.zid = sortid
	sorttype = sortArrayValue(2)
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

<script type="text/javascript" language="javascript">
function chgtype(){
	switch($('#sorttype').val()){
		case '1':
			//$('#moreInput').show();$('#moreText').hide();
			$('#sortlist').val('about.asp');
			$('#sortshow').val('about.asp');
			break;
		case '2':
			$('#sortlist').val('newslist.asp');
			$('#sortshow').val('newsshow.asp');
			break;
		case '3':
			$('#sortlist').val('prolist.asp');
			$('#sortshow').val('proshow.asp');
			break;
		case '4':
			$('#moreInput').show();$('#moreText').hide();
			$('#sortlist').val('请在此添入外部地址');
			$('#sortlist').click(function(){if($('#sortlist').val()=='请在此添入外部地址')$('#sortlist').val('');});
			$('#sortshow').val('--');
			break;
		case '5':
			$('#sortlist').val('msgbook.asp');
			$('#sortshow').val('--');
			break;
		default:
			$('#sortlist').val('');
			$('#sortshow').val('');
			break;
	}
}
function delsort(tid){
	if (confirm('删除栏目会删除该目录下所有新闻、单页、产品等内容,并且无法恢复!!\n你确认要删除吗?')){
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
</script>
</head>

<body>
	<div id="wrapper">
    	<h1><%=manageName%></h1>
        
        <!-- You can name the links with lowercase, they will be transformed to uppercase by CSS, we prefered to name them with uppercase to have the same effect with disabled stylesheet -->
        <ul id="mainNav">
        	<%=manageMenu(2,0)%>
        </ul>
        <!-- // #end mainNav -->
        
        <div id="containerHolder">
			<div id="container">
        		<div id="sidebar">
                	<ul class="sideNav">
                    	<%=smallMenu(2,2)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">栏目管理</a> &raquo; <a href="#" class="active"><%if isup = false then%>栏目列表<%else%>栏目修改<%end if%></a></h2>
                
                <div id="main">
                	
                	<form action="<%=actstr%>" method="post" target="act" class="jNice">
					
                        
                        <%if isup = false then%>
                        <h3>栏目列表</h3>
                    	<%response.Write(typeclass.typeList)%>
                        <h3></h3>
                        <button name="" type="button" onclick="$('#sortadd').show();"><span><span>添加新栏目</span></span><div></div></button>
                        
                        <%end if%>
                         <div class="clear" style="height:20px;"></div>
                      <div id="sortadd" <%if isup = false then response.Write("style=""display:none""") end if%>>  
                      <%if isup = false then%>
                      <h3><a name="add">添加栏目</a></h3>
                      <% else%>
				  		<h3><a name="add">修改栏目</a></h3>
				  		<%
				  		end if
				  		%>
                    	<fieldset>
                        	<p><label>栏目名称:</label><input name="sortname" type="text" class="text-long" value="<%=sortArrayValue(1)%>" /></p>
                        	<p><label>栏目类型:</label>
                            <select onchange="chgtype();" name="sorttype" id="sorttype">
                            	<%=sorttypelist(sorttype)%>
                            </select>
                            </p>
                            <p><label>父级栏目:</label>
                            <%
							'typeclass.fid = 6
							response.Write(typeclass.typeInput)%>
                    </p>
                    <p><label>栏目选项:</label>
                        <span>是否显示在主导航</span><input type="checkbox" <%if CBool(sortArrayValue(4)) = true then response.Write("checked=""checked""") end if%> name="SortIsNav" value="y" />
                        <span>是否显示在二级导航</span><input name="SortIsList" value="y" <%if CBool(sortArrayValue(5)) = true or isup=false then response.Write("checked=""checked""") end if%> type="checkbox" /></p>
                    <p><label>排序数字(越小越前):</label><input type="text" class="text-small" name="sortorder" value="<%if isup = true then response.write(sortArrayValue(6)) else response.write "10" end if%>" /></p>
                    <%if isup = true then %>
                    <div id="moreInput">
                    <%else%>
                    <div id="moreText"><p><a href="javascript:void(0);" onclick="$('#moreInput').show();$('#moreText').hide();">显示更多选项...</a></p></div>
                    <div id="moreInput" class="jNiceHidden">
                    <%end if%>
                        <p><label>关键词:</label><input type="text" name="sortkeywords" value="<%=sortArrayValue(7)%>" class="text-long" /></p>
                        <p><label>栏目说明:</label><input type="text" name="sortdescription" value="<%=sortArrayValue(8)%>" class="text-long" /></p>
                        <p><label>栏目图片<a href="javascript:;"  name="pic" class="thickboxs" rel="#hiddenModalContent" style="text-decoration:none;">上传图片</a>:</label>
                        <input type="text"  name="sortpic" id="pic" value="<%=sortArrayValue(9)%>" class="text-long" /></p>
                        
                        <p><label>栏目(列表)页:</label><input type="text" name="SortListTemplates" value="<%if isup = true then response.Write(sortArrayValue(10)) else response.Write("about.asp") end if%>" class="text-long" id="sortlist" /></p>
                        <p><label>栏目内容页:</label><input type="text" name="SortShowTemplates" value="<%if isup = true then response.Write(sortArrayValue(11)) else response.Write("about.asp") end if%>" class="text-long" id="sortshow" /></p>
                         <p><label>自定义字段(仅支持图片和产品,格式( name:中文名|name2:中文名2 )):</label><input type="text" name="SortDiyField" value="<%=sortArrayValue(13)%>" class="text-long" /></p>
                    </div>
                    
                    
                        	<p><label>栏目简介:</label>
                            
                            
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
							oFCKeditor.Value	= sortArrayValue(12)
							oFCKeditor.Create "sortcontent"
							%>
                            </p>
                            <%if isup = false then%>
                            <button type="submit"><span><span>添加新栏目</span></span><div></div></button>
                            <%else%>
                            <button type="submit"><span><span>修 改</span></span><div></div></button>
                            <button type="button" onclick="window.location='sort.asp';"><span><span>返 回</span></span><div></div></button>
                            <%end if%>
                        </fieldset>
                        </div>
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
	set typeclass = nothing
	dbclose
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
