<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="menu.asp"-->
<%
checklogin 100
Dim title,groups
title = "欢迎登陆益东集团网站内容管理系统"

'dbopen

%>
<%

Response.Buffer = true

' 声明待检测数组
Dim ObjTotest(26,4)

ObjTotest(0,0) = "MSWC.AdRotator"
ObjTotest(1,0) = "MSWC.BrowserType"
ObjTotest(2,0) = "MSWC.NextLink"
ObjTotest(3,0) = "MSWC.Tools"
ObjTotest(4,0) = "MSWC.Status"
ObjTotest(5,0) = "MSWC.Counters"
ObjTotest(6,0) = "IISSample.ContentRotator"
ObjTotest(7,0) = "IISSample.PageCounter"
ObjTotest(8,0) = "MSWC.PermissionChecker"
ObjTotest(9,0) = "Scripting.FileSystemObject"
	ObjTotest(9,1) = "(FSO 文本文件读写)"
ObjTotest(10,0) = "adodb.connection"
	ObjTotest(10,1) = "(ADO 数据对象)"
	
ObjTotest(11,0) = "SoftArtisans.FileUp"
	ObjTotest(11,1) = "(SA-FileUp 文件上传)"
ObjTotest(12,0) = "SoftArtisans.FileManager"
	ObjTotest(12,1) = "(SoftArtisans 文件管理)"
ObjTotest(13,0) = "LyfUpload.UploadFile"
	ObjTotest(13,1) = "(刘云峰的文件上传组件)"
ObjTotest(14,0) = "Persits.Upload.1"
	ObjTotest(14,1) = "(ASPUpload 文件上传)"
ObjTotest(15,0) = "w3.upload"
	ObjTotest(15,1) = "(Dimac 文件上传)"

ObjTotest(16,0) = "JMail.SmtpMail"
	ObjTotest(16,1) = "(Dimac JMail 邮件收发) <a href='http://www.ajiang.net'>中文手册下载</a>"
ObjTotest(17,0) = "CDONTS.NewMail"
	ObjTotest(17,1) = "(虚拟 SMTP 发信)"
ObjTotest(18,0) = "Persits.MailSender"
	ObjTotest(18,1) = "(ASPemail 发信)"
ObjTotest(19,0) = "SMTPsvg.Mailer"
	ObjTotest(19,1) = "(ASPmail 发信)"
ObjTotest(20,0) = "DkQmail.Qmail"
	ObjTotest(20,1) = "(dkQmail 发信)"
ObjTotest(21,0) = "Geocel.Mailer"
	ObjTotest(21,1) = "(Geocel 发信)"
ObjTotest(22,0) = "IISmail.Iismail.1"
	ObjTotest(22,1) = "(IISmail 发信)"
ObjTotest(23,0) = "SmtpMail.SmtpMail.1"
	ObjTotest(23,1) = "(SmtpMail 发信)"
	
ObjTotest(24,0) = "SoftArtisans.ImageGen"
	ObjTotest(24,1) = "(SA 的图像读写组件)"
ObjTotest(25,0) = "W3Image.Image"
	ObjTotest(25,1) = "(Dimac 的图像读写组件)"

public IsObj,VerObj,TestObj
public okOS,okCPUS,okCPU

'检查预查组件支持情况及版本

dim i
for i=0 to 25
	on error resume next
	IsObj=false
	VerObj=""
	'dim TestObj
	TestObj=""
	set TestObj=server.CreateObject(ObjTotest(i,0))
	If -2147221005 <> Err then		'感谢网友iAmFisher的宝贵建议
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if
	ObjTotest(i,2)=IsObj
	ObjTotest(i,3)=VerObj
next

'检查组件是否被支持及组件版本的子程序
sub ObjTest(strObj)
	on error resume next
	IsObj=false
	VerObj=""
	TestObj=""
	set TestObj=server.CreateObject (strObj)
	If -2147221005 <> Err then		'感谢网友iAmFisher的宝贵建议
		IsObj = True
		VerObj = TestObj.version
		if VerObj="" or isnull(VerObj) then VerObj=TestObj.about
	end if	
End sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=title%></title>
<link href="style/css/transdmin.css" rel="stylesheet" type="text/css" media="screen" />
<script type="text/javascript" src="../script/jquery132.js"></script>
<script type="text/javascript" src="style/js/jNice.js"></script>
<%=GetNewVersion%>
</head>

<body>
	<div id="wrapper">
    	<h1><%=manageName%></h1>
        
        <!-- You can name the links with lowercase, they will be transformed to uppercase by CSS, we prefered to name them with uppercase to have the same effect with disabled stylesheet -->
        <ul id="mainNav">
        	<%=manageMenu(0,0)%>
        </ul>
        <!-- // #end mainNav -->
        
        <div id="containerHolder">
			<div id="container">
        		<div id="sidebar">
                	<ul class="sideNav">
                    	<%=smallMenu(0,1)%>
                    </ul>
                    <!-- // .sideNav -->
					
                </div>    
                <!-- // #sidebar -->
				
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">管理首页</a> &raquo; <a href="#" class="active">Welcome</a></h2>
                
                <div id="main">

					<h3>
                    欢迎您登陆益东集团网站内容管理系统:今天是 <%=Application("date_chinese")%><br />
                    
                    </h3>
                    	<fieldset>
                        	<p>
                       	    <label>版本信息</label>
                            &nbsp;&nbsp;&nbsp;&nbsp;<span>系统版本: <font style="line-height:250%; color: #5494af;"><%=Versions%></font> </span>
                            </p>
                       	 		  
                          <p>
                       	    <label>服务器信息:</label>
							服务器名：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=Request.ServerVariables("SERVER_NAME")%></font>
							服务器IP：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=Request.ServerVariables("LOCAL_ADDR")%></font>
							服务器端口：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=Request.ServerVariables("SERVER_PORT")%></font>
							服务器时间：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=now%></font>
							IIS版本：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=Request.ServerVariables("SERVER_SOFTWARE")%></font>
							脚本超时时间：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=Server.ScriptTimeout%></font>
							本文件路径：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=Request.ServerVariables("PATH_TRANSLATED")%></font>
							服务器解译引擎：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion%></font>
							服务器CPU数量：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=okCPUS%>个</font>
							服务器CPU详情：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=okCPU%></font>
							服务器操作系统：<font style="line-height:250%; margin-left:15px;color: #5494af;"><%=okOS%></font>
                          </p>
                       <p><label> IIS自带的ASP组件</label></p>
					  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; width:100%; margin-bottom:20px;" bordercolor="#F0F0F0" >
						<tr height="18">
						  <td width="320" bgcolor="#666666">组 件 名 称</td>
						  <td width="130" bgcolor="#666666">支持及版本</td>
						</tr>
						<%For i=0 to 10%>
						<tr height="18" class="backq">
						  <td align="left" bgcolor="#F9F9F9">&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
						  <td align="left" bgcolor="#F9F9F9">&nbsp;
							  <%
							If Not ObjTotest(i,2) Then 
								Response.Write "<font color=red><b>×</b></font>"
							Else
								Response.Write "<font class=fonts><b>√</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
							End If
							%>
						</td>
						</tr>
						<%next%>
					  </table>

					   <p>
                       	    <label>开发者信息</label>
                       	    &nbsp;&nbsp;&nbsp;&nbsp;作者：<font style="line-height:250%; color: #5494af;">JAM</font>
                            Email：<font style="line-height:250%; color: #5494af;">design2014@gmail.com</font>
                            网站：<font style="line-height:250%; color: #5494af;"><a href="http://www.jamvs.com" target="_blank">http://www.jamvs.com</a></font>
                          </p>
                          <p>
                       	    <label>相关支持</label>
                            &nbsp;&nbsp;&nbsp;&nbsp;网络咨询：<font style="line-height:250%; color: #5494af;">QQ95149313</font>
                          </p>				
                          <div class="clear"></div>

                      </fieldset>
                      
                      
                      <div class="clear"></div>
                	
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
</body>
</html>
