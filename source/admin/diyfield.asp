<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/function.asp"-->
<%
'返回自定义字段
checklogin 10

dim sortid,fieldarray,fi,nid,conStr
dim valueArray(50)
'on error resume next
sortid = int(request.QueryString("sid"))
nid = request.QueryString("nid")
if nid <> "" then nid = int(nid)
'if err.number <> 0 or sortid = 0 then
'	err.clear
'	response.end()
'end if

dbopen
	eRs.open "select top 1 SortDiyField,sorttype from sorts where sortid = " & sortid,eConn,0,1
	if not eRs.eof then
		if int(eRs("sorttype")) = 3 then
			if nid <> "" then
				conStr = econn.execute("select top 1 proDiyValue from products where proid =" & nid)(0)
				conStr = split(inull(conStr),"$$$")
				for fi = 0 to ubound(conStr)
					if ubound(split(conStr(fi),"||")) > 0 then
					valueArray(fi) = split(conStr(fi),"||")(1)
					end if
				next
			end if
		end if
		if inull(eRs("SortDiyField")) <> "" then
			fieldarray = split(eRs("SortDiyField"),"|")
			for fi = 0 to ubound(fieldarray)
				if ubound(split(fieldarray(fi),":")) > 0 then
				response.Write("<p><label>"&split(fieldarray(fi),":")(1)&":</label><input type=""text"" value="""&valueArray(fi)&""" name=""diy_"&split(fieldarray(fi),":")(0)&""" class=""text-long"" /></p>")
				end if
			next
		end if
	end if
dbclose
%>