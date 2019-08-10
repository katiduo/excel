<%@ page language="java" contentType="text/html; charset=UTF-8"
	pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
    <link type="text/css" rel="stylesheet">
    	
    </link>
<body>
<form action="/importExcel" method="post" enctype="multipart/form-data">
	<input type="file" name="filePro" id="a_idPicPath" />
	<a href="${pageContext.request.contextPath }/download">下载</a>
	<a href="${pageContext.request.contextPath }/importExcel">导入</a>
	<a href="/excelExport">导出</a>
	${sessionScope.num}
	<input type="submit" value="提交"/>
</form>
</body>
</html>
