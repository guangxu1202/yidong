<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/typeclass.asp"-->
<!--#include file="menu.asp"-->

<%

checklogin 5
Dim title,groups,tagArray,i,tagid,isup,actstr
Dim intpage,pagecounts,pagesizes,recordcounts,newssql
dim tagArrayValue(50)
isup = false

intpage = request.QueryString("page")
if intpage = "" or not isnumeric(intpage) then intpage = 1
intpage = int(intpage)

if not isnumeric(tagid) then
	tagid = 0
else
	tagid = int(tagid)
end if


title = "标签样式管理 - 欢迎登陆网站内容管理系统"

tagArray = array("tagID","tagName","tagContent","tagtype","tagRemark","tagAttr")
if request.QueryString("act") = "add" then
	for i = 1 to ubound(tagArray)
		tagArrayValue(i) = request.Form(tagArray(i))
	next
	if tagArrayValue(1) = "" then writealert "样式名称不能为空",0
	dbopen
	eRs.open "select * from templateTags where tagname = '"&tagArrayValue(1)&"'",eConn,1,3
	if not eRs.eof then 
		dbclose
		 writealert "样式名称不能重复",0
	end if
	eRs.Addnew
	for i = 1 to ubound(tagArray)
		'response.Write(newsArray(i) & ":")
		eRs(tagArray(i)) = tagArrayValue(i)
	next
	ers("tagissys") = false
	eRs.update
	eRs.Close
	dbclose
	writealert "添加成功",1
elseif request.QueryString("act") = "del" then
	tagid = int(request.Form("id"))
	dbopen
	eConn.execute("delete from templateTags where tagid =  " & tagid)
	dbclose
	response.Write("0")
	response.End()
elseif request.QueryString("act") = "ups" then
	tagid = int(request.QueryString("id"))
	for i = 1 to ubound(tagArray)
		tagArrayValue(i) = request.Form(tagArray(i))
	next
	if tagArrayValue(1) = "" then writealert "样式名称不能为空",0
	dbopen
	eRs.open "select * from templateTags where tagname = '"&tagArrayValue(1)&"' and tagid <> " & tagid,eConn,1,3
	if not eRs.eof then
		dbclose
		writealert "样式名称有重复",0
	end if
	eRs.Close
	eRs.open "select * from templateTags where  tagid = " & tagid,eConn,1,3

	'eRs.Addnew
	for i = 1 to ubound(tagArray)
		response.Write(tagArray(i) & ":")
		eRs(tagArray(i)) = tagArrayValue(i)
	next
	eRs.update
	eRs.Close
	dbclose
	writealert "修改成功","tagstyle.asp?page="&intpage
end if


dbopen
Dim cls
set cls = new typeCls
cls.selectname = "tagsort"

if request.QueryString("act") = "up" then
	tagid = int(request.QueryString("id"))
	isup = true
	
	actstr = "?act=ups&page="&intpage&"&id=" & tagid
	eRs.open "select top 1 * from templateTags where tagid = " & tagid,eConn,0,1
		for i = 1 to ubound(tagArray)
			tagArrayValue(i) = eRs(tagArray(i))
		next
	eRs.Close
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
<script type="text/javascript" src="../script/jquery.caretInsert.js"></script>
<script type="text/javascript" src="style/js/jNice.js"></script>
<script type="text/javascript" src="style/js/chgtype.js"></script>
<script language="javascript" type="text/javascript">
function deltag(tid){
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
$(document).ready(
				  function($){
					  $('#tagcontent').setCaret();
					  chgtype();
					  $('#tagcontent').dblclick(function(){$(this).height($(this).height()+50);});
					  $('#tagRemark').dblclick(function(){$(this).height($(this).height()+50);});
					}				  
				  );
function insertStr(str){$('#tagcontent').insertAtCaret(str);}
</script>
</head>

<body>

	<div id="wrapper">
    	<h1><%=manageName%></h1>
        
        <!-- You can name the links with lowercase, they will be transformed to uppercase by CSS, we prefered to name them with uppercase to have the same effect with disabled stylesheet -->
        <ul id="mainNav">
        	<%=manageMenu(3,0)%>
        </ul>
        <!-- // #end mainNav -->
        
        <div id="containerHolder">
			<div id="container">
        		<div id="sidebar">
                	<ul class="sideNav">
                    	<%=smallMenu(1,4)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">模板管理</a> &raquo; <a href="#" class="active">标签样式管理</a></h2>
                
                <div id="main">
                	<%if isup = false then%>
                	<form action="" class="jNice" target="act" method="post">
					<h3>样式列表</h3>
                    	<table cellpadding="0" cellspacing="0">
                        <%
						

						pagesizes = 15						
						
						recordcounts = eConn.execute("select count(*) from templateTags")(0)
						if (recordcounts/pagesizes) > int(recordcounts/pagesizes) then
							pagecounts = int(recordcounts/pagesizes + 1)
						else
							pagecounts = int(recordcounts/pagesizes)
						end if
						if pagecounts < 1 then pagecounts = 1
						if intpage > pagecounts then intpage = pagecounts
						if intpage = 1 then 
							newssql = "select top "&pagesizes&" * from templateTags order by tagid desc"
						else
							newssql = "select top "&pagesizes&" * from templateTags where tagid not in (select top "&(intpage-1)*pagesizes&" tagid from templateTags  order by tagid desc) order by tagid desc"
						end if
						'response.Write(newssql)
						eRs.open newssql,eConn,0,1
						do while not eRs.eof
						i = i + 1
						%>
							<tr id="t<%=eRs("tagid")%>" <%if i mod 2 = 0 then response.Write("class=""odd""") end if%>>
                                <td style="width:10px;"><%=eRs("tagid")%></td><td>
								[<%=tagTypeName(eRs("tagtype"))%>] <%=eRs("tagname")%>
                                <%if eRs("tagissys") = true then response.Write("&nbsp;<span>[系统]</span>")%>
                                </td>
                                <td class="action"><a href="?act=up&page=<%=intpage%>&id=<%=eRs("tagid")%>" class="edit">修改</a><a href="javascript:void(0);" onclick="deltag(<%=eRs("tagid")%>);" class="delete">删除</a></td>
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
                        </form>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3><a name="add"></a>添加标签</h3>
                        <%else%>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3>修改标签</h3>
                        <%end if%>
                    	<fieldset>
                        	<p><label>标签名称:</label><input type="text" name="tagname" value="<%=tagArrayValue(1)%>" class="text-long" /></p>
                            <p><label>标签类别:</label>
				
                            <select onchange="chgtype();" name="tagType" id="tagType">
                            <%response.Write(tagTypeList(tagArrayValue(3)))%>
                            </select>

                            </p>
                            <!--div id="artattr" style="display:none;">
                                <p><label>选择栏目(当前栏目请不要选择):</label>
                                    <%response.Write(cls.typeInput)%>
                                </p>
                                <p><label>标题字数:</label><input type="text" name="tagnamexxxxxxxx" value="" class="text-long" /></p>
                                <p><label>调用条数:</label><input type="text" name="tagnamexxxxxxxxxxx" value="" class="text-long" /></p>
                              <p>
                                    <label>日期样式:</label>
                                    <select>
                                        <option>YYYY-MM-DD</option>
                                        <option>MM-DD</option>
                                        <option>YYYY/MM/DD</option>
                                        <option>MM/DD</option>
                                    </select>
                                </p>
                              <p><label>摘要字数(0为不显示):</label><input type="text" name="tagnamexxxxxxxxx" value="" class="text-long" /></p>
                                <p><label>其它:</label>
                                    包含子类:<input name="ArtIsTopxxxx"  type="checkbox" value="y" />
                                    显示分页:<input name="ArtIsTopxxxxx"  type="checkbox" value="y" />
                                    仅显示图片新闻:<input name="ArtIsTopxxxxxxx"  type="checkbox" value="y" />
                                </p>
                            </div-->
                            
                            <p><label>显示样式:[点击下面的字段可直接插入]</label></p>
                            <p id="taglist">
                            	<span class="tagspan" onclick="insertStr('{网站名称}');">网站名称</span>
                                <span class="tagspan" onclick="insertStr('{网站版权}');">网站版权</span>
                                <span class="tagspan" onclick="insertStr('{网站域名}');">网站域名</span>
                                <span class="tagspan" onclick="insertStr('{网站关键词}');">网站关键词</span>
                                <span class="tagspan" onclick="insertStr('{网站说明}');">网站说明</span>
                                <span class="tagspan" onclick="insertStr('{网站备案号}');">网站备案号</span>
                                <span class="tagspan" onclick="insertStr('{当前位置}');">当前位置</span>
                                <span class="tagspan" onclick="window.open('http://www.escms.cn/help.asp');">帮助</span>
                            </p>
                            <p id="taglistsub">
                            </p>
                            <p>
                            <textarea rows="1" name="tagcontent" id="tagcontent" cols="1"><%=replace(tagArrayValue(2),"</textarea","&lt;/textarea")%></textarea>
                            
                            </p>
                            <p>提示：双击输入框可放大输入框</p>
                            <p><label>样式说明:</label>
                            <textarea rows="1" name="tagRemark" id="tagRemark" cols="1"><%=replace(tagArrayValue(4),"</textarea","&lt;/textarea")%></textarea>
                            </p>

                        	<%if isup = false then%>
                            <button type="submit"><span><span>添 加</span></span><div></div></button>
                            <%else%>
                            <button type="submit"><span><span>修 改</span></span><div></div></button>
                            <button type="button" onclick="window.location='tagstyle.asp';"><span><span>返 回</span></span><div></div></button>
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
<iframe frameborder="0" style="display:none;" name="act"></iframe>
</body>
</html>
