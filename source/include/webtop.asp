<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="webtags.asp"-->
<!--#include file="md5.asp"-->
<%
dim diyfield,tSortType,tSortID
dim fieldArray(61) '0-10 栏目用  11-20  单页用  20-40 新闻用 40-60 产品用  60+其它
tSortType = ""
tSortID = 0
dbopen
Init
%>