<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/md5.asp"-->
<%
if request.QueryString("act") = "out" then
	response.Cookies(CookiesKey)("username") = ""
	response.Cookies(CookiesKey)("usergroup") = ""
	response.Cookies(CookiesKey)("usercode") = ""
end if
if request.QueryString("act") = "login" then
	
	dim username,userpass,getcode
	username = trim(request.Form("username"))
	userpass = trim(request.Form("userpass"))
	getcode = trim(request.Form("getcode"))
	if getcode <> Session("GetCode") then
		Session("GetCode") = ""
		response.Write("验证码错误!")
		response.End()
	end if
	if username = "" or  userpass = "" then
		response.Write("请添写完整!")
		response.End()
	end if
	dbopen
		eRs.open "select top 1 userid,username,userpass,usergroup from users where username = '"&replace(username,"'","")&"' and userpass = '"&EnPas(md5(userpass))&"' and usergroup < 11",eConn,0,1
		if eRs.eof then
			dbclose
			response.Write("用户名或密码错误!")
			response.End()
		else
			response.Cookies(CookiesKey)("username") = eRs("username")
			response.Cookies(CookiesKey)("usergroup") = eRs("usergroup")
			response.Cookies(CookiesKey)("usercode") = EnPas(eRs("username") & eRs("usergroup") & eRs("userpass"))
			response.Write("0")
		end if
	dbclose
	response.End()
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>网站内容管理系统-后台登陆</title>
<link href="style/css/transdmin.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../script/jquery132.js"></script>
<style type="text/css">
body {background-color:#e5f1fa;}
#login {margin:80px auto; width:873px; height:462px; background-image:url(style/img/login_bg.gif);}
.left {width:48%; float:left;}
.right {width:50%; float:right; margin-top:100px; border-left:1px #666 solid; height:300px;}
.left h1 {margin-left:100px; margin-top:200px; background-image:url(style/img/logo.gif);  background-repeat:no-repeat; width:259px; height:92px; text-indent:-9999px;}
.tips {text-align:left; margin-left:130px; margin-top:20px;}
.tips dt {font-size:18px;}
.tips ul {list-style-type:decimal; margin-left:40px; margin-top:10px; font-size:12px;}
.tips ul {line-height:24px;}
.tips a {text-decoration:none; color:#666;}

.loginform {margin-top:40px;text-align:left; padding-left:50px; position:relative;}
.loginform p {height:40px;}
.loginform label {width:50px; display:block; float:left; line-height:30px;}
.loginform h3 {font-size:14px; color:#036431; margin-bottom:30px; margin-left:0px!important; padding:0px;}
.loginform input {height:25px; background:url(style/img/inputbg.gif); border:1px #cdcdcd solid; line-height:25px; padding-left:5px; font-size:12px; font-family:"Courier New", Courier, monospace;}
.loginform input.long { width:150px;}
.loginform input.small { width:100px;}
.loginform img {vertical-align:middle;}
.loginform .error { background-color:#B00; width:210px; color:#fff; height:24px; line-height:24px; margin-bottom:10px; display:none;}
</style>
<script language="javascript" type="text/javascript">
function cklogin(){
	$('.error').html('&nbsp;&nbsp;正在登陆...!');
	$('.error').show(300);	
	if ($('#UserName').val()==''){
		$('#UserName').focus();
		$('.error').html('&nbsp;&nbsp;用户名不能为空!');
		return;
	}
	if ($('#UserPass').val()==''){
		$('#UserPass').focus();
		$('.error').html('&nbsp;&nbsp;密码不能为空!');
		return;
	}
	if ($('#GetCode').val()==''){
		$('#GetCode').focus();
		$('.error').html('&nbsp;&nbsp;验证码不能为空!');
		return;
	}
	$.ajax({type:'POST',
		   url:'login.asp?act=login',
		   data:'username='+escape($('#UserName').val())+'&userpass='+escape($('#UserPass').val())+'&getcode='+escape($('#GetCode').val()),
		   success:function(msg){
			   if(msg=='0'){window.location='default.asp';}
			   else{$('.error').html('&nbsp;&nbsp;'+msg);$('.error').show();}
			   if(msg=='验证码错误!'){$('#codeimg').attr('src','../include/getcode.asp?'+Math.random());}
		   }});
	//$('#codeimg').attr('src','../include/getcode.asp?'+Math.random());
}
function KeyDown()
{
    if (event.keyCode == 13)
    {
        event.returnValue=false;
        event.cancel = true;
        cklogin();
    }
}
</script>
</head>

<body>
<div id="login">
	<div class="left">
    	<h1>恒泰高尔夫网站管理系统</h1>
    </div>
    <div class="right">
    	<div class="loginform">
        	<h3>欢迎登陆网站内容管理系统</h3>
            <form>
            <p><label>用户名:</label><input type="text" name="UserName" id="UserName" class="long" /></p>
            <p><label>密&nbsp;&nbsp;&nbsp;&nbsp;码:</label><input type="password" class="long" name="UserPass" id="UserPass" /></p>
            <p><label>验证码:</label><input type="text" name="GetCode" class="small" id="GetCode" onkeydown="KeyDown()" />
            <img src="../include/getcode.asp?" border="0" id="codeimg" onclick="this.src+=Math.random();" /></p>
            <div class="error">&nbsp;&nbsp;用户名错误!</div>
            <p><a href="javascript:void(0);" onclick="cklogin();"><img src="style/img/login.gif" alt="登 陆" width="177" height="33" border="0" /></a></p>
            </form>
        </div>
    </div>
</div>
</body>
</html>
