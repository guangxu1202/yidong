<% 
'*************************** 
'名称：FSO操作类 ESCMS
'描述：FSO操作类 
'*************************** 
Class FsoCls
	Private Fso'//对象 
	Public FsoObj'//公共接口对象 
	
	'//初始化,构造函数 
	Private Sub Class_Initialize 
		Set Fso=CreateObject(FSOname) 
		Set FsoObj=Fso 
	End Sub 
	'//结束，释构函数 
	Private Sub Class_Terminate 
		Set Fso=Nothing 
		Set FsoObj=Nothing 
	End Sub 
	
	
	'====================文件操作开始========================= 
	Function IsFileExists(ByVal FileDir) 
	'判断文件是否存在，存在则返回True,否则返回False 
	'参数FileDir为文件的绝对路径 
		If Fso.FileExists(FileDir) Then 
			IsFileExists=True 
		Else 
			IsFileExists=False 
		End If 
	End Function 
	Function GetFileText(ByVal FileDir) 
	'读取文件内容，存在则返回文件的内容，否则返回False 
	'参数FileDir为文件的绝对路径 
		If IsFileExists(FileDir) Then 
			Dim F 
			Set F=Fso.OpenTextFile(FileDir) 
			GetFileText=F.ReadAll 
			Set F=Nothing 
		Else 
			GetFileText=False 
		End If 
	End Function 
	Function CreateFile(ByVal FileDir,ByVal FileContent) 
	'创建一个文件并写入内容 
	'操作成功返回True,否则返回False 
		If IsFileExists(FileDir) Then 
			CreateFile=False 
			Exit Function 
		Else 
			Dim F 
			Set F=Fso.CreateTextFile(FileDir) 
			F.Write FileContent 
			CreateFile=True 
			F.Close 
		End If 
	End Function 
	Function DelFile(ByVal FileDir) 
	'删除一个文件，成功返回True，否则返回False 
	'参数FileDir为文件的绝对路径 
		If IsFileExists(FileDir) Then 
			Fso.DeleteFile(FileDir) 
			DelFile=True 
		Else 
			DelFile=False 
		End If 
	End Function 
	
	Function CopyFile(ByVal FromFile,ByVal ToFile) 
	'删除一个文件，成功返回True，否则返回False 
	'参数FileDir为文件的绝对路径 
	Dim F
		If IsFileExists(FromFile) Then 
			Set F = Fso.GetFile(FromFile) 
			F.Copy ToFile,true
			CopyFile=True 
		Else 
			CopyFile=False 
		End If 
	End Function 
	
	
	'====================文件操作结束========================= 
	
	'====================文件夹操作开始======================== 
	Function IsFolderExists(ByVal FolderDir) 
	'判断文件夹是否存在，存在则返回True,否则返回False 
	'参数FolderDir为文件的绝对路径 
		If Fso.FolderExists(FolderDir) Then 
			IsFolderExists=True 
		Else 
			IsFolderExists=False 
		End If 
	End Function 
	Sub CreateFolderA(ByVal ParentFolderDir,ByVal NewFoldeName) 
	'//在某个特定的文件夹里创建一个文件夹 
	'//ParentFolderDir为父文件夹的绝对路径，NewFolderName为要新建的文件夹名称 
		If IsFolderExists(ParentFolderDir&"\"&NewFoldeName) Then Exit Sub 
		Dim F,F1 
		Set F=Fso.GetFolder(ParentFolderDir) 
		Set F1=F.SubFolders 
		F1.Add(NewFoldeName) 
		Set F=Nothing 
		Set F1=Nothing 
	End Sub 
	Sub CreateFolderB(ByVal NewFolderDir) 
	'//创建一个新文件夹 
	'//参数NewFolderDir为要创建的文件夹绝对路径 
		If IsFolderExists(NewFolderDir) Then Exit Sub 
		Fso.CreateFolder(NewFolderDir) 
	End Sub 
	Sub DeleteAFolder(ByVal NewFolderDir) 
	'//删除一个新文件夹 
	'//参数NewFolderDir为要创建的文件夹绝对路径 
		If IsFolderExists(NewFolderDir)=False Then 
			Exit Sub 
		Else 
			Fso.DeleteFolder (NewFolderDir) 
		End If 
	End Sub 
	Function FolderItem(ByVal FolderDir) 
	'//文件夹里的文件夹集合 
	'//FolderDir 为文件夹绝对路径 
		If IsFolderExists(FolderDir) =False Then 
			FolderItem=False 
			Exit Function 
		End If 
		Dim FolderObj,FolderList,F 
		Set FolderObj=Fso.GetFolder(FolderDir) 
		Set FolderList=FolderObj.SubFolders 
		FolderItem=FolderObj.SubFolders.Count'//文件夹总数 
		For Each F In FolderList 
			FolderItem=FolderItem&"|"&F.Name 
		Next 
		Set FolderList=Nothing 
		Set FolderObj=Nothing 
	End Function 
	
	Function FileItem(ByVal FolderDir) 
	'//文件夹里的文件集合 
	'//FolderDir 为文件夹绝对路径 
		If IsFolderExists(FolderDir) =False Then 
			FileItem=False 
			Exit Function 
		End If 
		Dim FileObj,FileerList,F 
		Set FileObj=Fso.GetFolder(FolderDir) 
		Set FileList=FileObj.Files 
		FileItem=FileObj.Files.Count'//文件总数 
		For Each F In FileList 
			FileItem=FileItem&"|"&F.Name 
		Next 
		Set FileList=Nothing 
		Set FileObj=Nothing 
	End Function 
End Class  


%>