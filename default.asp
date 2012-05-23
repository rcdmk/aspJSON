<!--#include file="json.class.asp" -->
<%
dim jsonObj, jsonString

testeLoad = true
teteAdd = false

jsonString = "{ ""chave"" : ""valorTexto"", ""chave2"": 123 }"

set jsonObj = new json

if testeLoad then
	jsonObj.load jsonString
end if

if teteAdd then
	dim arr, multArr
	arr = Array("teste", 234, "mais teste", "234")
	
	redim multArr(1, 1)
	multArr(0, 0) = "0,0"
	multArr(0, 1) = "0,1"
	multArr(1, 0) = "1,0"
	multArr(1, 1) = "1,1"
	
	
	jsonObj.add "nome", "Jozé"
	jsonObj.add "idade", 25
	jsonObj.add "lista", arr
	jsonObj.add "lista2", multArr
end if


%>
<pre>
	<%= jsonObj.write %>
</pre>