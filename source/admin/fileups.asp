<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<!--#include file="../include/fsocls.asp"-->
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
                    	<%=smallMenu(1,8)%>
                    </ul>
                    <!-- // .sideNav -->
                </div>    
                <!-- // #sidebar -->
                
                <!-- h2 stays for breadcrumbs -->
                <h2><a href="#">系统管理</a> &raquo; <a href="#" class="active">文件管理</a></h2>
                
                <div id="main">
                	
                	<form action="" class="jNice" target="act" method="post">
					<h3>文件列表</h3>
                    	<table cellpadding="0" cellspacing="0">
                        <%
											
						
						dim fs,FolderPath,FolderAllPath,FolderList,FileList,fi
						FolderPath = request.QueryString("ph")
						FolderPath = replace(FolderPath,"..","_")
						if FolderPath = "" then FolderPath = "/"
						
						FolderAllPath = FolderPath
						
						if Sysfolder = "" then
							FolderPath = "/userfiles" & FolderPath
						else
							FolderPath = "/" & Sysfolder & "/userfiles" & FolderPath
						end if
						'response.Write(FolderAllPath)
						set fs = new fsocls
						FolderList = split(fs.FolderItem(server.MapPath(FolderPath)),"|")
						%>
                        	<tr id="t<%=fi%>" <%if i mod 2 = 0 then response.Write("class=""odd""") end if%>>
                                <td style="width:10px;"></td><td>
								[ <a href="?ph=<%=server.URLEncode(mid(FolderAllPath,1,InStrRev(left(FolderAllPath,len(FolderAllPath)-1),"/")))%>">..</a> ]
                                </td>
                                <td class="action"><a href="?ph=<%=server.URLEncode(mid(FolderAllPath,1,InStrRev(left(FolderAllPath,len(FolderAllPath)-1),"/")))%>" class="delete">返回上级目录</a></td>
                            </tr> 
						<%
						for fi = 1 to ubound(FolderList)
						%>

							<tr id="f<%=fi%>" <%if fi mod 2 = 0 then response.Write("class=""odd""") end if%>>
                            <td style="width:10px;"></td>
                                <td>
								[ <a href="?ph=<%=server.URLEncode(FolderAllPath &  FolderList(fi) & "/")%>"><%=FolderList(fi)%></a> ]
                                </td>
                                <td class="action"><a href="javascript:void(0);" onclick="deltag('<%=FolderAllPath &  FolderList(fi)%>');" class="delete">删除</a></td>
                            </tr>  
                        <%
						next
						FileList = split(fs.FileItem(server.MapPath(FolderPath)),"|")
						pagesizes = 15	
						recordcounts = FileList(0)
						if (recordcounts/pagesizes) > int(recordcounts/pagesizes) then
							pagecounts = int(recordcounts/pagesizes + 1)
						else
							pagecounts = int(recordcounts/pagesizes)
						end if
						if pagecounts < 1 then pagecounts = 1
						
						dim arrpage
						arrpage = (intpage - 1) * pagesizes
						for fi = 1 to 15
						 if fi + arrpage > ubound(FileList) then exit for
						%>

							<tr id="t<%=fi + arrpage%>" <%if fi mod 2 = 0 then response.Write("class=""odd""") end if%>>
                                <td style="width:10px;"><input name="nid" type="checkbox" value="<%="/userfiles"&FolderAllPath &  FileList(fi + arrpage)%>" /></td><td>
								<%
								dim exts
								exts = right(FileList(fi + arrpage),3)
								if  exts = "jpg" or exts = "gif" or exts= "png" then response.Write("<img src=""/userfiles"&FolderAllPath &  FileList(fi + arrpage)&""" width=""30"" height=""20""  />") end if%> 
                                <a target="_blank" href="/userfiles<%=FolderAllPath &  FileList(fi + arrpage)%>"><%=FileList(fi + arrpage)%></a>
                                
                                </td>
                                <td class="action"><a href="javascript:void(0);" onclick="delfile('<%="/userfiles"&FolderAllPath &  FileList(fi + arrpage)%>');" class="delete">删除</a></td>
                            </tr>  
                        <%
						next
						%>                       
                        </table>
                        <div id="page">
						
						<%=WritePage(pagecounts,intpage)%>
                        </div>
                        
                        <div class="clear"></div>
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
