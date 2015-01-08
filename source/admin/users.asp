<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/md5.asp"-->
<!--#include file="menu.asp"-->

<%
checklogin 5
Dim title,groups,userArray,i,userid,isup,actstr
Dim intpage,pagecounts,pagesizes,recordcounts,newssql
dim userArrayValue(50)
isup = false

intpage = request.QueryString("page")
if intpage = "" or not isnumeric(intpage) then intpage = 1
intpage = int(intpage)




title = "用户管理 - 欢迎登陆网站内容管理系统"

userArray = array("UserID","UserName","UserPass","UserGroup","UserPower","UserPic","UserEmail","UserAddress","UserContent","UserTel")
if request.QueryString("act") = "add" then
	for i = 1 to ubound(userArray)
		userArrayValue(i) = request.Form(userArray(i))
	next
	if userArrayValue(1) = "" then writealert "用户名不能为空",0
	if userArrayValue(2) = "" then writealert "密码不能为空",0 'EnPas
	userArrayValue(2) = EnPas(md5(userArrayValue(2)))
	dbopen
	eRs.open "select * from users where 1<>1",eConn,1,3
	eRs.Addnew
	for i = 1 to ubound(userArray)
		'response.Write(newsArray(i) & ":")
		eRs(userArray(i)) = userArrayValue(i)
	next
	
	ers("UserRegTime") = now()
	eRs.update
	eRs.Close
	dbclose
	writealert "添加成功",1
elseif request.QueryString("act") = "del" then
	userid = int(request.Form("id"))
	dbopen
	eConn.execute("delete from users where userid =  " & userid)
	dbclose
	response.Write("0")
	response.End()
elseif request.QueryString("act") = "ups" then
	userid = int(request.QueryString("id"))
	for i = 1 to ubound(userArray)
		userArrayValue(i) = request.Form(userArray(i))
	next
	if userArrayValue(1) = "" then writealert "用户名不能为空",0
	if userArrayValue(2) = "" then writealert "密码不能为空",0
	dbopen
	eRs.open "select * from users where userid = " & userid,eConn,1,3
	'eRs.Addnew

	if eRs("UserPass") <> userArrayValue(2) then userArrayValue(2) = EnPas(md5(userArrayValue(2)))
	
	for i = 2 to ubound(userArray)
		'response.Write(userArray(i) & ":")
		eRs(userArray(i)) = userArrayValue(i)
	next
	eRs.update
	eRs.Close
	dbclose
	writealert "修改成功","users.asp?page="&intpage
end if


dbopen

if request.QueryString("act") = "up" then
	userid = int(request.QueryString("id"))
	isup = true
	
	actstr = "?act=ups&page="&intpage&"&id=" & userid
	eRs.open "select top 1 * from users where userid = " & userid,eConn,0,1
		for i = 1 to ubound(userArray)
			userArrayValue(i) = eRs(userArray(i))
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
<script type="text/javascript" src="style/js/jNice.js"></script>

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
                    	<%=smallMenu(1,5)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">系统管理</a> &raquo; <a href="#" class="active">用户管理</a></h2>
                
                <div id="main">
                	<%if isup = false then%>
                	<form action="" class="jNice" target="act" method="post">
					<h3>用户列表</h3>
                    	<table cellpadding="0" cellspacing="0">
                        <%
						

						pagesizes = 15						
						
						recordcounts = eConn.execute("select count(*) from users")(0)
						if (recordcounts/pagesizes) > int(recordcounts/pagesizes) then
							pagecounts = int(recordcounts/pagesizes + 1)
						else
							pagecounts = int(recordcounts/pagesizes)
						end if
						if pagecounts < 1 then pagecounts = 1
						if intpage > pagecounts then intpage = pagecounts
						if intpage = 1 then 
							newssql = "select top "&pagesizes&" * from users order by userid desc"
						else
							newssql = "select top "&pagesizes&" * from users where userid not in (select top "&(intpage-1)*pagesizes&" userid from users  order by userid desc) order by userid desc"
						end if
						'response.Write(newssql)
						eRs.open newssql,eConn,0,1
						do while not eRs.eof
						i = i + 1
						%>
							<tr id="t<%=eRs("userid")%>" <%if i mod 2 = 0 then response.Write("class=""odd""") end if%>>
                                <td>
								<%=eRs("username")%>
                                <%if eRs("userGroup") <= 10 then response.Write("&nbsp;<span>[管]</span>")%>
                                </td>
                                <td class="action"><a href="?act=up&page=<%=intpage%>&id=<%=eRs("userid")%>" class="edit">修改</a><a href="javascript:void(0);" onclick="deltag(<%=eRs("userid")%>);" class="delete">删除</a></td>
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
						<h3><a name="add"></a>新建用户</h3>
                        <%else%>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3>修改用户</h3>
                        <%end if%>
                    	<fieldset>
                        	<p><label>用户名:</label><input type="text" name="UserName" value="<%=userArrayValue(1)%>" class="text-long" /></p>
                            <p><label>密码:</label><input type="password" name="UserPass" value="<%=userArrayValue(2)%>" class="text-long" /></p>
                            <p><label>用户组:</label>
                            <select name="usergroup" id="usergroup">
                            <%
							
							response.Write(usergroupList(userArrayValue(3)))
							%>
                            </select>
                            </p>
                            <p><label>头像:</label><input type="text" name="UserPic" value="<%=userArrayValue(5)%>" class="text-long" /></p>
                            
                            <p><label>Email:</label><input type="text" name="UserEmail" value="<%=userArrayValue(6)%>" class="text-long" /></p>
                            <p><label>地址:</label><input type="text" name="UserAddress" value="<%=userArrayValue(7)%>" class="text-long" /></p>
                            <p><label>电话:</label><input type="text" name="UserTel" value="<%=userArrayValue(9)%>" class="text-long" /></p>
                            
                            <p><label>备注:</label><textarea rows="1" name="UserContent" cols="1"><%=userArrayValue(8)%></textarea>
                            </p>

                        	<%if isup = false then%>
                            <button type="submit"><span><span>添 加</span></span><div></div></button>
                            <%else%>
                            <button type="submit"><span><span>修 改</span></span><div></div></button>
                            <button type="button" onclick="window.location='users.asp';"><span><span>返 回</span></span><div></div></button>
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
