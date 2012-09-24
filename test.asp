<!--#include file="json.class.asp" -->
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>ASPJSON</title>
	
	<style type="text/css">
		body {
			font-family: Arial, Helvetica, sans-serif;
		}
	
		pre {
			border: solid 1px #CCCCCC;
			background-color: #EEE;
			padding: 5px;
			text-indent: 0;
			width: 90%;
			word-break: break-strict;
			word-wrap: break-word;
		}
	</style>
</head>
<body>
	<%
	server.ScriptTimeout = 10
	dim jsonObj, jsonString, jsonArr
	
	testLoad = true
	testAdd = false
	testValue = false
	testChange = false
	
	testArrayPush = true
	
	
	set jsonObj = new json
	set jsonArr = new jsonArray
	
	jsonObj.debug = true
	
	if testLoad then
		jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""arrays"": [1, ""2"", 3.4, [5, 6, [7, 8]]], ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] } }"
		
		jsonObj.parse jsonString
		%>
		<h3>Input</h3>
		<pre><%= jsonString %></pre>
		<%
	end if
	
	if testAdd then
		dim arr, multArr, nestedObject
		arr = Array(1, "teste", 234.56, "mais teste", "234")
		
		redim multArr(1, 1)
		multArr(0, 0) = "0,0"
		multArr(0, 1) = "0,1"
		multArr(1, 0) = "1,0"
		multArr(1, 1) = "1,1"
		
		
		jsonObj.add "nome", "Jozé"
		jsonObj.add "idade", 25
		jsonObj.add "lista", arr
		jsonObj.add "lista2", multArr
		
		set nestedObject = new JSON
		nestedObject.add "sub1", "value of sub1"
		nestedObject.add "sub2", "value of sub2"
		
		jsonObj.add "nested", nestedObject
	end if
	
	
	if testValue then
		%><h3>Get the Values</h3><%
		response.write "nome: " & jsonObj.value("nome") & "<br>"
		response.write "idade: " & jsonObj.value("idade") & "<br>"
	end if
	
	
	if testChange then
		%><h3>Change the Values</h3><%
		
		response.write "nome before: " & jsonObj.value("nome") & "<br>"
		
		jsonObj.change "nome", "Mario"
		
		response.write "nome after: " & jsonObj.value("nome") & "<br>"
		
		jsonObj.change "nonExisting", -1
		
		response.write "Non existing property is created with: " & jsonObj.value("nonExisting") & "<br>"
	end if
	
	%>
	<h3>Output</h3>
	<pre><%= jsonObj.write %></pre>	
	<%
	
	if testArrayPush then
		jsonArr.Push jsonObj
		jsonArr.Push 1
		jsonArr.Push "strings too"
		%>
		<h3>Array Output</h3>
		<pre><%= jsonArr.write %></pre>	
		<%
	end if
	
	
	set jsonObj = nothing
	set jsonArr = nothing
	%>
</body>
</html>
