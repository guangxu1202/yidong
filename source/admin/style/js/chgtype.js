var start;
function chgtype(){
	
	switch($('#tagType').val())
	{
	case "1":  //全局信息
		if(start)$(start).hide();
		$('#taglistsub').html('');
		break;
	case "2":    //导航通用
		if(start)$(start).hide();
		$('#taglistsub').html('');
		
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<ul>\\n[LIST]\\n<li {selected}><a href={连接地址}>{栏目名称}</a></li>\\n[/LIST]\\n</ul>\');">DIV默认样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table><tr>\\n[LIST]\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n[/LIST]\\n</tr></table>\');">TABLE横向默认样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table>\\n[LIST]<tr>\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n</tr>[/LIST]\\n</table>\');">TABLE竖向默认样式</span>');
		
		var tags = Array('栏目名称','连接地址','栏目ID','selected');
		for(var tag in tags){
			$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}

		break;
	case "3":  //栏目通用
		if(start)$(start).hide();
		$('#taglistsub').html('');
		var tags = Array('栏目名称','连接地址','栏目ID','栏目关键词','栏目说明','栏目图片');
		for(var tag in tags){
			$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}
		break;
	case "4":  //新闻列表
		if(start)$(start).hide();
		//$('#artattr').show();
		start = '#artattr';
		
		$('#taglistsub').html('');

		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<ul>\\n[LIST]\\n<li {selected}><a href={连接地址}>{栏目名称}</a></li>\\n[/LIST]\\n</ul>\');">DIV默认样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table><tr>\\n[LIST]\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n[/LIST]\\n</tr></table>\');">TABLE默认样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table><tr>\\n[LIST]\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n[/LIST]\\n</tr></table>\');">TABLE两列样式</span>');
		$('#taglistsub').append('<br class="clear" />')

		var tags = Array('栏目名称','栏目地址','栏目ID','LIST','/LIST','|','文章ID','文章地址','标题','更新时间','文章点击数','来源','作者','摘要','图片地址');
		for(var tag in tags){
			if (tags[tag]=='|') $('#taglistsub').append('<br class="clear" />')
			else if(tags[tag]=='LIST' ||tags[tag]=='/LIST' ) $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'['+tags[tag]+']\');">'+tags[tag]+'</span>');
			else $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}

		break;
	case "5":  //产品列表
		if(start)$(start).hide();
		$('#taglistsub').html('');

		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<ul>\\n[LIST]\\n<li {selected}><a href={连接地址}>{栏目名称}</a></li>\\n[/LIST]\\n</ul>\');">DIV图片样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<ul>\\n[LIST]\\n<li {selected}><a href={连接地址}>{栏目名称}</a></li>\\n[/LIST]\\n</ul>\');">DIV文本样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table><tr>\\n[LIST]\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n[/LIST]\\n</tr></table>\');">TABLE图片样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table><tr>\\n[LIST]\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n[/LIST]\\n</tr></table>\');">TABLE文本样式</span>');
		$('#taglistsub').append('<br class="clear" />')

		var tags = Array('栏目名称','栏目地址','栏目ID','LIST','/LIST','|','产品ID','产品点击数','产品地址','产品名称','更新时间','产品小图地址','产品大图地址','摘要','价格','车辆','机票','高尔夫','餐饮','酒店','其他','包含','相关介绍','注意事项');
		for(var tag in tags){
			if (tags[tag]=='|') $('#taglistsub').append('<br class="clear" />')
			else if(tags[tag]=='LIST' ||tags[tag]=='/LIST' ) $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'['+tags[tag]+']\');">'+tags[tag]+'</span>');
			else $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{diy:字段名称}\');">自定义字段</span>');

		break;
	case "6":  //留言列表
		if(start)$(start).hide();
		$('#taglistsub').html('');

		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<ul>\\n[LIST]\\n<li {selected}><a href={连接地址}>{栏目名称}</a></li>\\n[/LIST]\\n</ul>\');">DIV图片样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<ul>\\n[LIST]\\n<li {selected}><a href={连接地址}>{栏目名称}</a></li>\\n[/LIST]\\n</ul>\');">DIV文本样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table><tr>\\n[LIST]\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n[/LIST]\\n</tr></table>\');">TABLE图片样式</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="$(\'#tagcontent\').val(\'<table><tr>\\n[LIST]\\n<td {selected}><a href={连接地址}>{栏目名称}</a></td>\\n[/LIST]\\n</tr></table>\');">TABLE文本样式</span>');
		$('#taglistsub').append('<br class="clear" />')

		var tags = Array('栏目名称','栏目地址','栏目ID','|','姓名','电话','email','留言内容','留言IP','留言时间');
		for(var tag in tags){
			if (tags[tag]=='|') $('#taglistsub').append('<br class="clear" />')
			else $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}
		//$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{diy:字段名称}\');">自定义字段</span>');

		break;
	case "7":  //单页  单页标题,单页ID,单页地址,单页图片,单页摘要,单页内容
		$('#taglistsub').html('');
		var tags = Array('栏目名称','栏目ID','|','单页标题','单页ID','单页地址','单页图片','单页摘要','单页内容');
		for(var tag in tags){
			if (tags[tag]=='|') $('#taglistsub').append('<br class="clear" />')
			else $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}
		break;
	case "8":  //新闻
		$('#taglistsub').html('');
		var tags = Array('栏目名称','栏目ID','|','文章ID','文章地址','标题','更新时间','文章点击数','来源','作者','摘要','图片地址','文章内容');
		for(var tag in tags){
			if (tags[tag]=='|') $('#taglistsub').append('<br class="clear" />')
			else $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}
		break;
	case "9":  //产品
		$('#taglistsub').html('');
		var tags = Array('栏目名称','栏目ID','|','产品ID','产品点击数','产品地址','产品名称','更新时间','产品小图地址','产品大图地址','摘要','产品内容');
		for(var tag in tags){
			if (tags[tag]=='|') $('#taglistsub').append('<br class="clear" />')
			else $('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{'+tags[tag]+'}\');">'+tags[tag]+'</span>');
		}
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'{diy:字段名称}\');">自定义字段</span>');

		break;
	case "10":  //留言框样式
		$('#taglistsub').html('');
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'<form onsubmit=&#34;return false;&#34; name=&#34;message&#34;>\\n\');">form</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'\\n<lable>用户名</lable\\n><input type=&#34;text&#34; name=&#34;username&#34; id=&#34;username&#34; /><span id=&#34;userTip&#34;></span>\');">用户名</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'\\n<lable>Email</lable>\\n<input type=&#34;text&#34; name=&#34;email&#34; id=&#34;email&#34; value=&#34;&#34; /><span id=&#34;emailTip&#34;></span>\');">EMAIL</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'\\n<lable>电话</lable\\n><input type=&#34;text&#34; name=&#34;phone&#34; id=&#34;phone&#34; /><span id=&#34;phoneTip&#34;></span>\');">电话</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'\\n<lable>内容</lable\\n><textarea rows=&#34;10&#34; cols=&#34;40&#34; name=&#34;content&#34; id=&#34;content&#34; ></textarea><span id=&#34;contentTip&#34;></span>\');">内容</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'\\n<input type=&#34;image&#34; onclick=&#34;msgsubmit()&#34; src=&#34;images/send.gif&#34;/>\');">提交</span>');
		$('#taglistsub').append('<span class="tagspan" onclick="insertStr(\'\\n</form>\');">/form</span>');
		break;
	case "11":  //html
		$('#taglistsub').html('');
		break;
	default:
		$('#taglistsub').html('');
		break;
	}
	$('#tagcontent').focus();
}