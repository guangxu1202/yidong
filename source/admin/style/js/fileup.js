var picid = 'pic';
function SendFile()
{
	var date = new Date();
	document.getElementById('frmUpload').action = '../fckeditor/editor/filemanager/connectors/asp/upload.asp?CurrentFolder=%2Fimage%2F&Type=Image&time=' + date.getTime();
	document.getElementById('frmUpload').submit() ;
}

function OnUploadCompleted( errorNumber, fileUrl, fileName, customMsg )
{
	switch ( errorNumber )
	{
		case 0 :
			jQuery('#'+picid).val(fileUrl);
			alert( '文件上传成功' ) ;
			closeAll();
			break ;
		case 1 :
			alert( customMsg ) ;
			break ;
		case 10 :
			jQuery('#'+picid).val(fileUrl);
			alert( customMsg ) ;
			closeAll();
			break ;
		case 201 :
			jQuery('#'+picid).val(fileUrl);
			alert( '文件上传成功,但文件名重复,已重命名为: "' + fileName + '"' ) ;
			closeAll();
			break ;
		case 202 :
			alert( '无效文件' ) ;
			break ;
		case 203 :
			alert( "文件写入错误,请检查服务器是否有权限上传文件!" ) ;
			break ;
		default :
			alert( '文件上传错误,错误代号: ' + errorNumber ) ;
			break ;
	}
}
function closeAll() {$(".thickboxs[rel]").each(function() {$(this).overlay().close();}); }
function InsertPic(picurl)
{
	var oEditor = FCKeditorAPI.GetInstance('content') ;
	if ( oEditor.EditMode == FCK_EDITMODE_WYSIWYG )
	{
		oEditor.InsertHtml( '<img src="'+picurl+'" border="0" />' ) ;
	}
	else
		alert( 'You must be on WYSIWYG mode!' ) ;
}
jQuery(function($){
$(".thickboxs[rel]").each(
  function(i){
	  $(this).overlay({
	  expose: {color: '#333',loadSpeed: 200,opacity: 0.9},
	  closeOnClick: false,
	  onBeforeLoad:function(){
		  var t=this.getTrigger().attr("name");
		  if (t!='') picid = t ;
		  else picid = 'pic';
	  }
	  });
  });
});