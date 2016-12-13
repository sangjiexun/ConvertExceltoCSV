<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Input Error</title>
</head>

<body>
<h2 style="color: red">Input Error</h2>

<%
String errorMsg = "";
//String errorMsg = request.getAttribute("Input").toString();
if ((request.getAttribute("InputError").toString()).equals("nullString")) {
	errorMsg = "ディレクトリが指定されていなません。ディレクトリを指定してください。";
}else {
	errorMsg = "何かしらのエラー。";
}
%>

<div style="color: red">
<%= errorMsg %>
</div>


</body>

</html>