<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/typeclass.asp"-->
<!--#include file="../fckeditor/fckeditor.asp" -->
<!--#include file="menu.asp"-->
<%
checklogin 10
Dim title,groups,aloneArray,i,sortid,aloneid,isup,actstr
dim aloneArrayValue(50)
title = "单页管理 - 欢迎登陆网站内容管理系统"
isup = false
aloneArray = array("AloneID","AloneName","AlonePic","AloneSort","AloneContent","AloneAbstract","AloneOrder","AloneShowPage")

if request.QueryString("act") = "add" then
	for i = 1 to ubound(aloneArray)
		aloneArrayValue(i) = request.Form(aloneArray(i))
	next
	if aloneArrayValue(1) = "" then writealert "名称不能为空",0
	dbopen
	eRs.open "select * from alones where 1<>1",eConn,1,3
	eRs.Addnew
	for i = 1 to ubound(aloneArray)
		'response.Write(aloneArray(i) & ":")
		eRs(aloneArray(i)) = repreg(aloneArrayValue(i))
	next
	eRs("AloneAddTime") = now()
	eRs.update
	eRs.Close
	dbclose
	writealert "添加成功",1
elseif request.QueryString("act") = "del" then
	AloneID = int(request.Form("id"))
	dbopen
	'response.Write("select top 1 aloneid from alones where aloneissort is not null and aloneid = " & aloneid)

	eRs.open "select top 1 aloneid from alones a where aloneissort is not null and exists(select 1 from sorts where sortid = a.aloneissort and sorttype = 1) and aloneid = " & aloneid,eConn,0,1
	if not eRs.eof then
		dbclose
		response.Write("此单页是一个栏目,不能删除.")
		response.End()
	end if
	eRs.Close
	eConn.execute("delete from alones where aloneid = " & AloneID)
	dbclose
	response.Write("0")
	response.End()
elseif request.QueryString("act") = "ups" then
	AloneID = int(request.QueryString("id"))
	for i = 1 to ubound(aloneArray)
		aloneArrayValue(i) = request.Form(aloneArray(i))
	next
	if aloneArrayValue(1) = "" then writealert "标题不能为空",0
	dbopen
	eRs.open "select * from alones where aloneid = " & AloneID,eConn,1,3
	'eRs.Addnew
	for i = 1 to ubound(aloneArray)
		'response.Write(aloneArray(i) & ":")
		eRs(aloneArray(i)) = repreg(aloneArrayValue(i))
	next
	
	eRs.update
	eRs.Close
	dbclose
	writealert "修改成功","alone.asp"
end if


dbopen
Dim cls
set cls = new typeCls
cls.selectname = "AloneSort"

if request.QueryString("act") = "up" then
	AloneID = int(request.QueryString("id"))
	isup = true
	
	actstr = "?act=ups&id=" & AloneID
	eRs.open "select top 1 * from alones where AloneID = " & AloneID,eConn,0,1
		for i = 1 to ubound(aloneArray)
			aloneArrayValue(i) = unrepreg(eRs(aloneArray(i)))
		next
	eRs.Close
	cls.fid = int(aloneArrayValue(3))
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
function delalone(tid){
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
                    	<%=smallMenu(3,1)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">网站内容管理</a> &raquo; <a href="#" class="active">单页管理</a></h2>
                
                <div id="main">
                	<form action="<%=actstr%>" method="post" target="act" class="jNice">
					<h3>单页列表</h3>
                    	<table cellpadding="0" cellspacing="0">
                        <%
						eRs.open "select * from alones order by aloneorder asc",eConn,0,1
						do while not eRs.eof 
						i = i + 1
						%>
							<tr id="t<%=eRs("aloneid")%>" <%if i mod 2 = 0 then response.Write(" class=""odd""") end if%>>
                                <td><%=eRs("aloneID")%></td><td><%=unrepreg(eRs("aloneName"))%></td>
                                <td class="action"><a href="../about.asp?aloneid=<%=eRs("aloneID")%>" target="_blank" class="view">查看</a><a href="?act=up&id=<%=eRs("aloneid")%>" class="edit">修改</a><a href="javascript:void(0);" onclick="delalone(<%=eRs("aloneid")%>);" class="delete">删除</a></td>
                            </tr>    
                         <%
						 eRs.movenext
						 loop
						 eRs.close
						 %>                    
						                    
                        </table>
						<h3><%if isup = true then response.Write("修改单页") else response.Write("添加单页") end if%></h3>
                    	<fieldset>
                        	<p><label>单页名称:</label><input name="AloneName" value="<%=aloneArrayValue(1)%>" type="text" class="text-long" /></p>
                            <p><label>导航图片:<a href="javascript:;"  name="pic" class="thickboxs" rel="#hiddenModalContent" style="text-decoration:none;">上传图片</a></label>
                            <input type="text" name="AlonePic" id="pic" value="<%=aloneArrayValue(2)%>" class="text-long" /></p>
                            <p><label>所属栏目:</label>
                            <%
							
							response.Write(cls.typeInput)
							%>
                            </p>
                          <p><label>显示页面:</label><input name="AloneShowPage" type="text" value="<%if isup = true then response.Write(aloneArrayValue(7)) else response.Write("about.asp") end if%>" class="text-long" /></p>
                          <p><label>单页排序:</label><input name="AloneOrder" type="text" class="text-medium" value="<%if isup = true then response.Write(aloneArrayValue(6)) else response.Write("10") end if%>" /></p>
                            <p><label>内容摘要:</label><textarea rows="1" name="AloneAbstract" cols="1"><%=aloneArrayValue(5)%></textarea></p>
                            <p><label>单页内容:</label>
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
							oFCKeditor.Width = 584
							oFCKeditor.Height = 400
							oFckeditor.ToolbarSet = "ESCMS"
							oFCKeditor.Value	= aloneArrayValue(4)
							oFCKeditor.Create "AloneContent"
							%>
                            </p>
                        	<%if isup = false then%>
                            <button type="submit"><span><span>添 加</span></span><div></div></button>
                            <%else%>
                            <button type="submit"><span><span>修 改</span></span><div></div></button><button type="button" onclick="window.location='alone.asp';"><span><span>返 回</span></span><div></div></button>
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
