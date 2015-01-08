<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<%
Dim msg,sortid,username,email,phone,content
sortid = request.Form("sortid")
username = repreg(request.Form("username"))
email = repreg(request.Form("email"))
phone = repreg(request.Form("phone"))
content = repreg(request.Form("content"))

if username = "" or username = "undefined" then
	response.Write("用户名不能为空!")
	response.End()
end if
if content = "" or content="undefined" then
	response.Write("内容不能为空")
	response.End()
end if
if not isnumeric(sortid) then
	sortid = ""
end if

content = replace(content,"script","_script_")
content = replace(content,"iframe","_iframe_")
content = RemoveHTML(content)
username = left(RemoveHTML(username),30)
email = left(RemoveHTML(email),50)
phone = left(RemoveHTML(phone),20)

if session("ismsg") <> "esmsg" then
	response.Write("提交不正确,请刷新页面重新提交,如果该问题重复存在,请联系技术支持!")
	response.End()
end if
if session("content") = Enpas(username & content) then
	response.Write("请勿重复提交!")
	response.End()
end if
session("content") = Enpas(username & content)
dbopen
eRs.open "select top 1 * from messages where 1=2",eConn,1,3
	eRs.addnew
	eRs("MsgTel") = phone
	eRs("MsgContent") = content
	eRs("MsgName") = username
	if sortid <> "" then
		eRs("MsgSort") = sortid
	end if
	eRs("MsgEmail") = email
	eRs("MsgAddTime") = now()
	eRs("MsgType") = "Msg"
	eRs("MsgIP") = userip
	if Application("web_IsMsgAudit") = true then
		eRs("MsgIsAudit") = false
	else
		eRs("MsgIsAudit") = true
	end if
	eRs.update
	eRs.Close
dbclose 
response.Write("0")



%>