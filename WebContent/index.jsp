<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Insert title here</title>
<script type="text/javascript" src="js/jquery-1.12.4.js"></script>
</head>
<body>
	<div>
	<input type="file" id="upfile" name="upfile" onchange="fileUpload();"/>
	<button type="button" name="btn">上传excel</button>
	<button onclick="downloadTemp();">下载模板</button>
	</div>
</body>
<script type="text/javascript">
	function downloadTemp(){
		window.location.href="/4sinfo/downloadTmpl.do";
	}
		
	function fileUpload(){
		var fileName = $("#upfile").val();
		if(fileName == null || fileName==""){
			alert("请选择文件");
		}else{
			var fileType = fileName.substr(fileName.length-4,fileName.length);
		if(fileType == ".xls" || fileType == "xlsx"){
			 var formData = new FormData();
			 formData.append("file",$("#upfile").prop("files")[0]);
			 $.ajax({
			 	type:"post",
			 	url:"/4sinfo/ajaxUpload.do",
			 	data:formData,
			 	cache:false,
			 	processData:false,
			 	contentType:false,
			 	dataType:"json",
			 	success:function(data){
			 		if(null != data){
			 			if(data.dataStatus == "1"){
			 				if(confirm("上传成功！")){
			 					window.location.reload();
			 				}
			 			}else{
			 				alert("上传失败！");
			 			}
			 		}
			 	},
			 	error:function(){
			 		alert("上传失败！");
			 	}
			 });
		}else{
			alert("上传文件类型错误！");
		}
		}
	}

</script>
</html>