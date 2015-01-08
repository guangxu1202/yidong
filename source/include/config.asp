<%
' ------------------------------------------


	' 以下各行不用修改
	Option Explicit
	Server.ScriptTimeout	= 5000
	Session.CodePage		= 65001
	Response.Charset		= "utf-8"
	Response.ContentType	= "text/html; charset=utf-8"
	Response.Buffer			= True 
	

	' 定义虚拟路径 <重要> 如果不会设置,请参考:http://www.escms.cn/help.asp
	Const Sysfolder = "yidong"
	' 定义CookiesKey <重要> 建议改成任意英文
	Const CookiesKey = "HTgolfkey"
	
	' 定义数据库类型,ACCESS/Sql
	Const DBType = "ACCESS"
	' ACCESS
	'定义数据库目录,请注意,仅仅为目录名称,不需要其它的!
	Const DBfolder = "database_yd"
	'定义数据库名称
	Const DBname = "JAMdb.mdb"
	' MSSql
	'数据库连接字符串
	Const SQLconnStr = "Provider=SQLOLEDB.1;Server=localhost;UID=sa;PWD=;Database=JAMdb"
	' FSO组件名称
	Const FSOname = "Scripting.FileSystemObject"


%>
