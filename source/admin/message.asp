<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/typeclass.asp"-->
<!--#include file="../fckeditor/fckeditor.asp" -->
<!--#include file="menu.asp"-->

<%
checklogin 10
Dim title,groups,msgArray,i,sortid,msgid,isup,actstr
Dim intpage,pagecounts,pagesizes,recordcounts,msgsql
dim msgArrayValue(50)
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


title = "留言管理 - 欢迎登陆网站内容管理系统"

msgArray = array("MsgID","MsgIP","MsgSort","MsgAddTime","MsgName","MsgTel","MsgEmail","MsgContent","MsgIsAudit","MsgReply")
if request.QueryString("act") = "del" then
	msgid = int(request.Form("id"))
	dbopen
	eConn.execute("delete from Messages where msgid = " & msgid)
	dbclose
	response.Write("0")
	response.End()
elseif request.QueryString("act") = "plup" then
	dim plid
	plid = request.Form("nid")
	plid = replace(plid,")","")
	if plid = "" then writealert "没有选择需要处理的项目！",0
	
	dbopen
	select case request.Form("acts")
	case "audit"
		eConn.execute("update messages set msgisaudit = true where msgid in("&plid&")")
	case "noaudit"
		eConn.execute("update messages set msgisaudit = false where msgid in("&plid&")")
	case "del"
		eConn.execute("delete from messages where msgid in("&plid&")")
	end select
	dbclose
	writealert "批量处理成功！","Message.asp?page="&intpage&"&sortid="&sortid
elseif request.QueryString("act") = "ups" then
	msgid = int(request.QueryString("id"))
	for i = 4 to ubound(msgArray)
		msgArrayValue(i) = repreg(request.Form(msgArray(i)))
	next

	if msgArrayValue(8) = "y" then msgArrayValue(8) = true else msgArrayValue(8) = false end if
	dbopen
	eRs.open "select * from Messages where msgid = " & msgid,eConn,1,3
	'eRs.Addnew
	for i = 4 to ubound(msgArray)
		'response.Write(msgArray(i) & ":")
		eRs(msgArray(i)) = msgArrayValue(i)
	next

	eRs.update
	eRs.Close
	dbclose
	writealert "修改或审核成功","Message.asp?page="&intpage&"&sortid="&sortid
end if


dbopen
Dim cls
set cls = new typeCls
cls.selectname = "msgSort"

if request.QueryString("act") = "up" then
	msgid = int(request.QueryString("id"))
	isup = true
	
	actstr = "?act=ups&page="&intpage&"&sortid="&sortid&"&id=" & msgid
	eRs.open "select top 1 * from Messages where msgid = " & msgid,eConn,0,1
		for i = 1 to ubound(msgArray)
			msgArrayValue(i) = unrepreg(eRs(msgArray(i)))
		next
	eRs.Close
	cls.fid = msgArrayValue(2)
	'typeclass.zid = sortid
	'sorttype = sortArrayValue(2)
else
	actstr = "?"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%></title>
<link href="style/css/transdmin.css" rel="stylesheet" type="text/css" media="screen" />
<link href="style/css/thickbox.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../script/jquery132.js"></script>
<script type="text/javascript" src="style/js/jNice.js"></script>
<script language="javascript" src="style/js/thickbox-compressed.js"></script>
<script language="javascript" src="style/js/fileup.js"></script>
<script language="javascript" type="text/javascript">
function tpic(){
	
}
function GetContents()
{
	var oEditor = FCKeditorAPI.GetInstance('ArtContent') ;
	return oEditor.GetXHTML( true )  ;
}
function delmsg(tid){
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
                    	<%=smallMenu(3,4)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">网站内容管理</a> &raquo; <a href="#" class="active">留言管理</a></h2>
                
                <div id="main">
                	<%if isup = false then%>
                	<form action="?act=plup&page=<%=intpage%>&sortid=<%=sortid%>" id="sbform" class="jNice" target="act" method="post">
					<h3>留言列表</h3>
                    	<p style=" margin:5px 10px; line-height:24px;">
                        
                        <a href="?">全部</a>&nbsp;&nbsp;
                        <%
						eRs.open "select sortid,sortname from sorts where sortid in (select msgSort from Messages)",eConn,0,1
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
							msgsql = " and msgsort = " & sortid
						end if
						pagesizes = 15						
						
						recordcounts = eConn.execute("select count(*) from Messages where 1=1 " & msgsql)(0)
						if (recordcounts/pagesizes) > int(recordcounts/pagesizes) then
							pagecounts = int(recordcounts/pagesizes + 1)
						else
							pagecounts = int(recordcounts/pagesizes)
						end if
						if pagecounts < 1 then pagecounts = 1
						if intpage > pagecounts then intpage = pagecounts
						if intpage = 1 then 
							msgsql = "select top "&pagesizes&" a.*,sortname from (Messages a left join sorts b on a.msgsort = b.sortid) where 1=1  " & msgsql & " order by MsgAddtime desc"
						else
							msgsql = "select top "&pagesizes&" a.*,sortname from (Messages a left join sorts b on a.msgsort = b.sortid) where msgid not in (select top "&(intpage-1)*pagesizes&" MsgId from Messages where 1=1 "&msgsql&" order by msgaddtime desc) " & msgsql & " order by msgaddtime desc"
						end if
						'response.Write(msgsql)
						eRs.open msgsql,eConn,0,1
						do while not eRs.eof
						i = i + 1
						%>
							<tr id="t<%=eRs("msgid")%>" <%if i mod 2 = 0 then response.Write("class=""odd""") end if%>>
                                <td style="width:10px;"><input name="nid" type="checkbox" value="<%=eRs("msgid")%>" /></td>
								<td>
								[<%=eRs("sortname")%>] <%=eRs("msgname")%> <%if eRs("msgIsAudit") = false then response.write("&nbsp;&nbsp;<span>[未审核]</span>") end if%>
                                </td>
								<td><%=getSubString(unrepreg(eRs("msgcontent")),10)%></td><td><%=eRs("msgemail")%></td><td><a href="http://www.ip138.com/ips.asp?ip=<%=eRs("msgip")%>" target="_blank"><%=eRs("msgip")%></a></td><td><%=Format_Time(eRs("msgaddtime"),6)%></td>
                                <td class="action"><a href="?act=up&sortid=<%=sortid%>&page=<%=intpage%>&id=<%=eRs("msgid")%>" class="edit">修改</a><a href="javascript:void(0);" onclick="delmsg(<%=eRs("msgid")%>);" class="delete">删除</a></td>
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
                       <button name="" type="button" onclick="plup('audit');"><span><span>批量审核</span></span><div></div></button>
                       <button name="" type="button" onclick="plup('noaudit');"><span><span>取消审核</span></span><div></div></button>
                        <button name="" type="button" onclick="plup('del');"><span><span>批量删除</span></span><div></div></button>
                        
                        <div class="clear"></div>
						
                        </form>
						<p>&nbsp;</p>
                        <%else%>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3>回复或审核留言</h3>
                        
                    	<fieldset>
                        	<p><label>姓名:</label><input type="text" name="MsgName" value="<%=msgArrayValue(4)%>" class="text-long" /></p>

                            <p><label>电话:</label>
                            <input type="text" name="MsgTel" id="MsgTel" value="<%=msgArrayValue(5)%>" class="text-long" /><br />
                            </p>
                            
                            <p><label>Email:</label><input type="text" name="MsgEmail" id="MsgEmail" value="<%=msgArrayValue(6)%>" class="text-long" />
                            </p>
							<p><label>审核:</label>
                            <span>审核通过</span><input name="MsgIsAudit" <%if CBool(msgArrayValue(8)) = true then response.Write("checked=""checked""")%> type="checkbox" value="y" />
                           </p>
						   <p><label>用户IP:</label></p>
						   <p>
						   <a href="http://www.ip138.com/ips.asp?ip=<%=msgArrayValue(1)%>" target="_blank"><%=msgArrayValue(1)%></a>
                            </p>

                             <p><label>留言内容:</label><textarea rows="1" name="MsgContent" cols="1"><%=msgArrayValue(7)%></textarea></p>
                            <p><label>回复:</label>
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
							oFCKeditor.Value	=msgArrayValue(9)
							oFCKeditor.Create "MsgReply"
							%>
                            </p>

                            <button type="submit"><span><span>修 改</span></span><div></div></button>
                            <button type="button" onclick="window.location='message.asp';"><span><span>返 回</span></span><div></div></button>

                        </fieldset>
						
                    </form>
					<%end if%>
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

    <form id="frmUpload" target="UploadWindow" enctype="multipart/form-data" action="" method="post">
    上传文件:<br>
    <p>
    <input type="file" name="NewFile" id="files" ><br /></p>
    <p>
    <button type="button" value="" onClick="SendFile();"><span><span>上 传</span></span><div></div></button>
    <button type="button" value="" onClick="tb_remove();"><span><span>取 消</span></span><div></div></button>
    </p>
    </form>
</div>
<iframe name="UploadWindow" width="0" height="0" style="display:none;"></iframe>
<iframe frameborder="0" style="display:none;" name="act"></iframe>
</body>
</html>
