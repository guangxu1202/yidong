<%
' ESCMS V2.X 通用无限分类 类
' 递归写法，效率不高。。建议类别数量在50以下，不然影响效率
' author: EF,usd@msn.cn
Class typeCls
	private fid_,zid_,selectname_,click_
	private sInputStr,Separator,selected,twotd
	Public Property Let fid(f_id)
		fid_ = f_id
	End Property
	public Property Let zid(z_id)
		zid_ = z_id
	End Property
	public Property Let selectname(names)
		selectname_ = names
	End Property
	public Property Let click(clicks)
		click_ = clicks
	End Property
	
	'类初始化
	Private Sub Class_Initialize 
		fid_ = 0
		zid_ = 0
		selectname_ = "sortfid"
		click_ = ""
	End Sub
	
	Private Sub Class_Terminate 
	End Sub
	
	Public Function typeList
		Dim InputStr
		Separator = ""
		sInputStr = ""
		InputStr = "<table cellpadding=""0"" cellspacing=""0"">" & vbCrLf
		SubTypeList "",0
		InputStr = InputStr & sInputStr
		InputStr = InputStr & "</table>" & vbCrLf
		typeList = InputStr
	End Function

	Private Function SubTypeList(Separator,sFid)
		Dim sRs,temptwo
		Set sRs = eConn.execute("select sortid,sortname,sortfid,sorttype from [sorts] where sortfid = " & sFid & " order by sortorder asc,sortid asc")
		
		do while not sRs.eof
			twotd = twotd + 1
			if twotd mod 2 = 0 then temptwo = " class=""odd""" else temptwo = "" end if
			if sRs("sortid") = fid_ then selected = "selected" else selected = "" end if
			sInputStr = sInputStr & "<tr "&temptwo&" id=""t"&sRs("sortid")&"""><td width=""10"">"&sRs("sortid")&"</td><td>" & Separator & sRs("sortname") & "<span> ["&sorttypename(sRs("sorttype"))&"]</span></td><td class=""action"">" & vbCrLf &_
						"<a href=""?act=up&id="&sRs("sortid")&""" class=""edit"">修改</a>" & vbCrLf &_
						"<a href=""javascript:void(0);"" onclick=""delsort("&sRs("sortid")&");"" class=""delete"">删除</a>" & vbCrLf &_
						"</td></tr>" & vbCrLf
			SubTypeList Separator & "├─",sRs("sortid")
			sRs.Movenext
		loop
		sRs.Close
		Set sRs = nothing
	End Function
	
	Public Function typeInput
		Dim InputStr
		Separator = ""
		sInputStr = ""
		InputStr = "<select name="""&selectname_&""" "&click_&" >" & vbCrLf
		InputStr = InputStr & "<option value=""0"">" & "根目录" & "</option>" & vbCrLf
		SubTypeInput "├─",0
		InputStr = InputStr & sInputStr
		InputStr = InputStr & "</select>" & vbCrLf
		typeInput = InputStr
	End Function
	
	Private Function SubTypeInput(Separator,sFid)
		Dim sRs
		Set sRs = eConn.execute("select sortid,sortname,sortfid from [sorts] where sortfid = " & sFid & " order by sortorder asc,sortid asc")
		do while not sRs.eof
			
			  if sRs("sortid") = fid_ then selected = "selected" else selected = "" end if
			  if sRs("sortid") <> zid_ then
			  	sInputStr = sInputStr & "<option " & selected & " value=""" & sRs("sortid") & """>" & Separator & sRs("sortname") & "</option>" & vbCrLf
			  	SubTypeInput Separator & "├─",sRs("sortid")
			  end if
			sRs.Movenext
		loop
		sRs.Close
		Set sRs = nothing
	End Function

End Class

%>