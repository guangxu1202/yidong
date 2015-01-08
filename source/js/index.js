$(window).load(function() {

	$('.banner').flexslider({
	    animation: 'slide',
	    pauseOnAction: false,
	    before: function(slider) {
	    	var target = slider.slides.eq(slider.animatingTo);
	    	TweenMax.fromTo(target.find('.text'), 2, {left: '95%'}, {left: '50%'});
	    	TweenMax.fromTo(target.find('.para'), 1.6, {left: '30%'}, {left: '0'});

		},
	    slideshowSpeed: 7000,           //Integer: Set the speed of the slideshow cycling, in milliseconds
	    animationSpeed: 1000,
	    animationLoop: false
	});
	
	$('.banner img').mousedown(function() {
		return false;
	});
	
	$('.newsslide').flexslider({
	    animation: 'slide',
	    controlNav: false,
	    slideshowSpeed: 5000
	});
	
	
	var kicks = $('.kwicks .kwick');
	var curIndex = 0;
	var arr = ["#218DCD",'#FFB110',"#FF721E","#3343AA","#00B862"];
	var tarNum =0;
	kicks.bind('mouseenter', function() {
		var tarIndex = kicks.index(this);
		if (curIndex != tarIndex) {
			tarNum =tarIndex;
			TweenMax.to(kicks.eq(curIndex),0.5, {width:115});
			TweenMax.to(kicks.eq(curIndex).find('.icon'),0.3, {backgroundColor: 'transparent'});
			TweenMax.to(kicks.eq(tarIndex),0.5, {width:535});
			TweenMax.to(kicks.eq(tarIndex).find('.icon'),0.3, {backgroundColor:arr[tarNum]});
			curIndex = tarIndex;
		}
	});
	
	function resizeBanner() {
		if (document.documentElement.clientWidth>1600) {
			$('.banner').removeClass('banner-s');
			$('.banner .para').css('margin-left', 0);
		} else {
			$('.banner').addClass('banner-s');		
			$('.banner .para').css('margin-left', document.documentElement.clientWidth/2-800);
		}
	}
	resizeBanner();
	$(window).resize(function() {
		resizeBanner();	
	});
	
});