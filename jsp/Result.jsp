<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.Iterator" %>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Result</title>
</head>

<body>
<h2>Result</h2>

<h3>Input Excel File</h3>
<%
ArrayList<String> arrayInFile = (ArrayList<String>)request.getAttribute("arrayInFile");
for(Iterator<String> iter = arrayInFile.iterator(); iter.hasNext(); ){
    out.println(iter.next() + "<br/>");
}
%>

<h3>Skip Excel File</h3>
<%
ArrayList<String> arraySkipFile = (ArrayList<String>)request.getAttribute("arraySkipFile");
for(Iterator<String> iter = arraySkipFile.iterator(); iter.hasNext(); ){
    out.println(iter.next() + "<br/>");
}
%>



</body>
</html>