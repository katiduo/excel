$(function() {
	//下载
	$("#J_download").click(function(){
		this.href="/excelExport"
	});

	$("#H_download").click(function(){
		this.href="/modelOut?name="+$(".name").val();
	})

	/*$('#J_download').on('click', function() {
		/!*$.ajax({
			url: "/excelExport",
			type: 'post',
			dataType: 'json',
			success: function (response) {
				if (response.code == 1) {
					window.open(response.download);
				} else {
					alert(response.message);
				}
			}
		});*!/
		location.href="/excelExport";
	});*/
	function time(){
		var start=setInterval(function(){
			$.ajax({
				url:"/index",
				type:"post",
				success:function(data){
					$("#content").html(Math.round(data))
					if(data==100){
						clearInterval(start);
					}
				}
			})
		},10);
	}
	//上传文件
	var uploader = WebUploader.create({
		auto: false,
		server: '/importExcel',
		pick: '#J_upload',
		resize: false,
		accept: {
			mimeTypes: '.xlsx,.xls,.png'
		}
	});

	uploader.on('fileQueued', function(file) {
		console.log('上传中...');
	});

	uploader.on('uploadSuccess', function(file, response) {
		console.log('上传成功');
		console.log(response.listData);
	});

	uploader.on('uploadError', function(file) {
		alert('上传出错');
	});


	$("#up").click(function(){
		time();
		uploader.upload();

	})



	setTimeout(function(){
		$('input').hide();
		$(".name").show();
	},20);
});