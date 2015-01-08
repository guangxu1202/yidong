<!--#include file="../../../../../include/config.asp" -->
<%
 ' fckeditor\editor\filemanager\connectors\asp
 ' FCKeditor - The text editor for Internet - http://www.fckeditor.net
 ' Copyright (C) 2003-2009 Frederico Caldeira Knabben
 '
 ' == BEGIN LICENSE ==
 '
 ' Licensed under the terms of any of the following licenses at your
 ' choice:
 '
 '  - GNU General Public License Version 2 or later (the "GPL")
 '    http://www.gnu.org/licenses/gpl.html
 '
 '  - GNU Lesser General Public License Version 2.1 or later (the "LGPL")
 '    http://www.gnu.org/licenses/lgpl.html
 '
 '  - Mozilla Public License Version 1.1 or later (the "MPL")
 '    http://www.mozilla.org/MPL/MPL-1.1.html
 '
 ' == END LICENSE ==
 '
 ' Configuration file for the File Manager Connector for ASP.
%>
<%

' SECURITY: You must explicitly enable this "connector" (set it to "True").
' WARNING: don't just set "ConfigIsEnabled = true", you must be sure that only
'		authenticated users can access this file or use some kind of session checking.
Dim ConfigIsEnabled
ConfigIsEnabled = True

' Path to user files relative to the document root.
' This setting is preserved only for backward compatibility.
' You should look at the settings for each resource type to get the full potential
Dim ConfigUserFilesPath
if Sysfolder <> "" then
	ConfigUserFilesPath = "/" & Sysfolder & "/userfiles/"
else
	ConfigUserFilesPath = "/userfiles/"
end if
' Due to security issues with Apache modules, it is recommended to leave the
' following setting enabled.
Dim ConfigForceSingleExtension
ConfigForceSingleExtension = true

' What the user can do with this connector
Dim ConfigAllowedCommands
ConfigAllowedCommands = "QuickUpload|FileUpload|GetFolders|GetFoldersAndFiles|CreateFolder"

' Allowed Resource Types
Dim ConfigAllowedTypes
ConfigAllowedTypes = "File|Image|Flash|Media"

' For security, HTML is allowed in the first Kb of data for files having the
' following extensions only.
Dim ConfigHtmlExtensions
ConfigHtmlExtensions = "html|htm|xml|xsd|txt|js"

Dim ConfigAllowedExtensions, ConfigDeniedExtensions, ConfigFileTypesPath, ConfigFileTypesAbsolutePath, ConfigQuickUploadPath, ConfigQuickUploadAbsolutePath
Set ConfigAllowedExtensions	= CreateObject( "Scripting.Dictionary" )
Set ConfigDeniedExtensions	= CreateObject( "Scripting.Dictionary" )
Set ConfigFileTypesPath	= CreateObject( "Scripting.Dictionary" )
Set ConfigFileTypesAbsolutePath	= CreateObject( "Scripting.Dictionary" )
Set ConfigQuickUploadPath	= CreateObject( "Scripting.Dictionary" )
Set ConfigQuickUploadAbsolutePath	= CreateObject( "Scripting.Dictionary" )

ConfigAllowedExtensions.Add	"File", "7z|aiff|asf|avi|bmp|csv|doc|fla|flv|gif|gz|gzip|jpeg|jpg|mid|mov|mp3|mp4|mpc|mpeg|mpg|ods|odt|pdf|png|ppt|pxd|qt|ram|rar|rm|rmi|rmvb|rtf|sdc|sitd|swf|sxc|sxw|tar|tgz|tif|tiff|txt|vsd|wav|wma|wmv|xls|xml|zip"
ConfigDeniedExtensions.Add	"File", ""
ConfigFileTypesPath.Add "File", ConfigUserFilesPath & "file/"
ConfigFileTypesAbsolutePath.Add "File", ""
ConfigQuickUploadPath.Add "File", ConfigUserFilesPath
ConfigQuickUploadAbsolutePath.Add "File", ""

ConfigAllowedExtensions.Add	"Image", "bmp|gif|jpeg|jpg|png"
ConfigDeniedExtensions.Add	"Image", ""
ConfigFileTypesPath.Add "Image", ConfigUserFilesPath & "image/"
ConfigFileTypesAbsolutePath.Add "Image", ""
ConfigQuickUploadPath.Add "Image", ConfigUserFilesPath
ConfigQuickUploadAbsolutePath.Add "Image", ""

ConfigAllowedExtensions.Add	"Flash", "swf|flv"
ConfigDeniedExtensions.Add	"Flash", ""
ConfigFileTypesPath.Add "Flash", ConfigUserFilesPath & "flash/"
ConfigFileTypesAbsolutePath.Add "Flash", ""
ConfigQuickUploadPath.Add "Flash", ConfigUserFilesPath
ConfigQuickUploadAbsolutePath.Add "Flash", ""

ConfigAllowedExtensions.Add	"Media", "aiff|asf|avi|bmp|fla|flv|gif|jpeg|jpg|mid|mov|mp3|mp4|mpc|mpeg|mpg|png|qt|ram|rm|rmi|rmvb|swf|tif|tiff|wav|wma|wmv"
ConfigDeniedExtensions.Add	"Media", ""
ConfigFileTypesPath.Add "Media", ConfigUserFilesPath & "media/"
ConfigFileTypesAbsolutePath.Add "Media", ""
ConfigQuickUploadPath.Add "Media", ConfigUserFilesPath
ConfigQuickUploadAbsolutePath.Add "Media", ""

%>
