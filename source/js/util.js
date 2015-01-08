$(function(){
	vNav();
	vSelect();
	vHon();
	vJust();
	vMap();
	vJob();
	vNews();

	vJobInput();
})


function vNav(){

	var speed = 300;
	$('.nav > li').hover(function(){
		$(this).children('a').addClass('active').next().stop().show();
	},function(){
		$(this).children('a').removeClass('active').next().stop().hide();
	})
}

function vSelect(){
	$('#footer-ol').children('li').each(function(){
		var drop = $(this).children('.drop');
		var a = $(this).children('a');
		var height = drop.height();
		drop.css('top',-height+'px').hide();

		$(this).hover(function() {
			drop.show();
			a.addClass('on');
		}, function() {
			drop.hide();
			a.removeClass('on');
		});
	})
}

function vMap(){
	var hd = $('.industry-map');
	var site_li = $('.siteUl li');
	var site_name = $('.siteUl li .site_name');
	var bd = $('.industry .industry-wrap').children();
	/*site_name.hover(function() {
		var hoverPic = $(this).parent().attr('hoverPic');
		$(this).parent().addClass(hoverPic)
	}, function() {
		var hoverPic = $(this).parent().attr('hoverPic');
		$(this).parent().removeClass(hoverPic)
	});*/
	if(bd.length>1){
		bd.hide().eq(15).show();
	}
	
	site_li.eq(15).addClass(site_li.eq(15).attr('hoverPic'));
	site_li.eq(15).find('.site_name').css("color",'#fff');
	site_name.bind('click',function(e){
		e.preventDefault();
		var index = $('.siteUl li').index($(this).parent());
		var hoverPic = $(this).parent().attr('hoverPic');
		site_li.each(function() {
			var pic = $(this).attr('hoverPic');
			$(this).removeClass(pic);
			$(this).find('.site_name').css("color",'#1e1e1e');
		});
		$(this).parent().addClass(hoverPic);
		$(this).css("color",'#fff');
		bd.hide().eq(index).show();
		$(document).scrollTop(hd.offset().top+$('.main').offset().top+200);
	});
}
function vNews(){
	var _this = $('.news-leftcon .in-slide');
	$('.prev-btn',_this).attr('href','javascript:void(0);');
	$('.next-btn',_this).attr('href','javascript:void(0);');
	_this.slide({
		mainCell:'ul',
		prevCell:'.prev-btn',
		nextCell:'.next-btn',
		effect:'fold'
	});
}
function vJust(){
	var tab = $('.bannerTab');
	var left = tab.children('.leftBtn');
	var right = tab.children('.rightBtn');
	var length = tab.find('li').length;
	if(length <= 1){
		left.hide();
		right.hide();
	}else{
		left.attr('href','javascript:void(0);');
		right.attr('href','javascript:void(0);');
		tab.slide({
			mainCell:'ul',
			effect:'fold',
			prevCell:left,
			nextCell:right
		})
	}

	$('.temp-banner').slide({
		titCell:'.hd span'
	});
	$('.temp-other-nr,.temp-slide-nr').slide({
		targetCell:'.trigger p'
	});
	$('.joinTab').slide({
		titCell:'.tabIcon li',
		mainCell:'.tabMian'
	})

	var img = 	$("#jcImg");
	if(img.length<1){return false};
	$("#jcImg").jcImgScroll({
		width :516, 
    	height:306,
    	title:false,
    	offsetX:75,
    	setZoom:.8
	});
}

function vHon(){
	var hon = $('.honour');
	var top_time = hon.find('.top-time');
	var bottom = hon.children('.bottom');
	var events = $('.events-bottom').children();
	var that = $('.honour-slide');
	var ul = $('.hd-nr ul',that);
	var width = ul.children('li').eq(0).outerWidth(true);
	var bool = true;
	ul.children('li').hover(function(){bool=false},function(){bool=true})
	that.slide({
		titCell:'.hd .hd-nr ul li',
		targetCell:'.trigger p',
		startFun:function(i){
			if(bool == true){
				var off = i-4>0?i-4:0;
				ul.animate({left:'-'+width*off+'px'},300)
			}
		}
	});
	top_time.find('li').bind('click',function(e){
		e.preventDefault();
		var index = $(this).index();
		$(this).addClass('on').siblings().removeClass('on');
		bottom.children().eq(index).show().siblings().hide();
		events.eq(index).show().siblings().hide();
		$('.temp-trval-trigger').children('div').eq(index).show().siblings().hide();
	}).eq(0).trigger('click');

	bottom.children('.temp-city').find('a').bind('click',function(e){
		e.preventDefault();
		$(this).addClass('on').parent().siblings().children('a').removeClass('on');
		var index = $(this).parent().index();
		var num = $(this).parents('.temp-city').index();

		$('.temp-trval-trigger').children('div').eq(num).children('div').eq(index).show().siblings().hide();
	})

	top_time.slide({
		mainCell:'.top-time-nr ul',
		effect:'left',
		vis:6,
		autoPage:true
	});

	$('.temp-trval-trigger').children('div').each(function(){
		$(this).children('div:not(:first)').hide();
	})

}
function vJob(){
	$(document).on('click',"tr",function() {
		$(this).next('.job-details').find(".details-wrap").stop().slideToggle();
	});
	$(document).on('click',".details-wrap .close",function() {
		$(this).parent().stop().slideUp();
	});
}

function vJobInput()
{
	$('#keywordSheHui').focus(function(){$(this).val('')});
	$('#keywordSchool').focus(function(){$(this).val('')});
}	