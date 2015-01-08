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
ProArray = array("ID","tips1","tips2")

if request.QueryString("act") = "ups" then
	newsid = 2
	for i = 1 to ubound(ProArray)
		ProArrayValue(i) = repreg(request.Form(ProArray(i)))
	next

	DBopen
	eRs.open "select *  from Tips where ID = " & newsid,eConn,1,3

	'eRs.Addnew

	for i = 1 to ubound(ProArray)
		'response.Write(ProArray(i) & ":")
		eRs(ProArray(i)) = ProArrayValue(i)
	next
	'eRs("ProUpdateTime") = now()
	'eRs("proDiyValue") = diyValue
	eRs.update
	eRs.Close
	dbclose
	writealert "修改成功","tips.asp?page="&intpage&"&sortid="&sortid
Else
	newsid = 2
	DBopen
	eRs.open "select *  from Tips where ID = " & newsid,eConn,1,1
	for i = 1 to ubound(ProArray)
		 ProArrayValue(i) = eRs(ProArray(i)) 
	next
	eRs.Close
	dbclose

end if


dbopen
Dim cls
set cls = new typeCls
cls.selectname = "ProSort"

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
                    	<%=smallMenu(3,8)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">网站内容管理</a> &raquo; <a href="#" class="active">提示框修改</a></h2>
                
                <div id="main">
                	<%if isup = false then%>
                	<form action="?act=ups&page=<%=intpage%>&sortid=<%=sortid%>" class="jNice" id="sbform" target="act" method="post">
                    	
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3><a name="add"></a>信息修改</h3>
                        <%else%>
                        <form action="<%=actstr%>" class="jNice" target="act" method="post">
						<h3>修改产品</h3>
                        <%end if%>
                    	<fieldset>
							
							<p><label>信息提示1：</label><textarea rows="1" name="tips1" cols="1"><%=ProArrayValue(1)%></textarea></p>
							<p><label>信息提示2：</label><textarea rows="1" name="tips2" cols="1"><%=ProArrayValue(2)%></textarea></p>


							
                            <button type="submit"><span><span>修 改</span></span><div></div></button>
                            <button type="button" onclick="window.location='hotel.asp';"><span><span>返 回</span></span><div></div></button>
                     
                      
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
</body>
</html>
