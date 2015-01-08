<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="menu.asp"-->
<%
checklogin 5
Dim title,groups
title = "系统设置 - 网站内容管理系统"
if request("act") = "up" then
	dim webhost,webname,webkeywords,webdescription,webbeian,webuser,IsMsgAudit,ishtml,webcopyright,errstr
	webname = request.Form("webname")
	webhost = request.Form("webhost")
	webkeywords = request.Form("webkeywords")
	webdescription = request.Form("webdescription")
	webbeian = request.Form("webbeian")
	webuser = request.Form("webuser")
	IsMsgAudit = request.Form("IsMsgAudit")
	ishtml = request.Form("ishtml")
	webcopyright = request.Form("webcopyright")
	if webname = "" then
		errstr = "网站名称不能为空！"
		writealert errstr,0
	end if
	
	if IsMsgAudit = "y" then IsMsgAudit = true else IsMsgAudit = false end if
	if ishtml = "y" then ishtml = true else ishtml = false end if
	
	dbopen
		eRs.open "select top 1 * from web_config",eConn,1,3
		eRs("webname") = webname
		eRs("webhost") = webhost
		eRs("webkeywords") = webkeywords
		eRs("webdescription") = webdescription 
		eRs("webbeian") = webbeian
		eRs("webuser") = webuser
		eRs("IsMsgAudit") = IsMsgAudit
		eRs("ishtml") = ishtml
		eRs("webcopyright") = webcopyright
		eRs.update
		eRs.Close
	dbclose
	BrandNewDay 1
	writealert "修改成功!",0
end if

dbopen

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%></title>
<link href="style/css/transdmin.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript" src="../script/jquery132.js"></script>
<script type="text/javascript" src="style/js/jNice.js"></script>
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
                    	<%=smallMenu(1,1)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">系统管理</a> &raquo; <a href="#" class="active">系统设置</a></h2>
                
                <div id="main">
                	<form action="?act=up" method="post" target="act" class="jNice">
				  <h3></h3>
                    	<fieldset>
                        	<p><label>网站名称:</label><input name="webname" value="<%=application("web_name")%>" type="text" class="text-long" /></p>
                            <p><label>网站域名:</label><input name="webhost" value="<%=application("web_host")%>" type="text" class="text-long" /></p>
                            <p><label>关键词:</label><input name="webkeywords" value="<%=application("web_keywords")%>" type="text" class="text-long" /></p>
                            <p><label>网站说明:</label><input name="webdescription" value="<%=application("web_description")%>" type="text" class="text-long" /></p>
                            <p><label>备案号:</label><input name="webbeian" value="<%=application("web_beian")%>" type="text" class="text-long" /></p>
                            <p><label>管理员名称:</label><input name="webuser" value="<%=application("web_user")%>" type="text" class="text-long" /></p>
                        <p><label>其它选项:</label>
                            	<span>留言需审核</span><input name="IsMsgAudit" type="checkbox" value="y" <%if Application("web_IsMsgAudit") = true then response.Write("checked=""checked""") end if%> />
                            <!--span>是否生成html</span><input name="ishtml" type="checkbox" <%if Application("web_ishtml") = true then response.Write("checked=""checked""") end if%> value="y" /-->
                          </p>
                        	<p><label>版权信息:</label><textarea name="webcopyright" rows="1" cols="1"><%=application("web_copyright")%></textarea></p>
                            <input type="submit" value="保 存" />
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
    <%dbclose%>
    <!-- // #wrapper -->
<iframe frameborder="0" style="display:none;" name="act"></iframe>
</body>
</html>
