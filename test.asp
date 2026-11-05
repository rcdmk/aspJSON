<%
Option Explicit
Response.LCID = 1046 ' Brazilian LCID (use your locale code here).
' Could also be the LCID property of the page declaration or Session.LCID to set it to the entire session.
Response.Buffer = true
%>
<!--#include file="jsonObject.class.asp" -->
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>ASPJSON Tests</title>
	<style type="text/css">
		body { font-family: Arial, Helvetica, sans-serif; }
		pre {
			border: solid 1px #CCCCCC;
			background-color: #EEE;
			padding: 5px;
			text-indent: 0;
			width: 90%;
			white-space: pre-wrap;
			word-wrap: break-word;
		}
		.pass { color: #0a0; }
		.fail { color: #c00; font-weight: bold; }
	</style>
</head>
<body>
	<h1>JSON Object and Array Tests</h1>
	<%
	Server.ScriptTimeout = 10

	dim passCount, failCount
	passCount = 0
	failCount = 0

	sub checkResult(cond, testName)
		if cond then
			passCount = passCount + 1
			response.write "<span class=""pass"">PASS</span>: " & testName & "<br>"
		else
			failCount = failCount + 1
			response.write "<span class=""fail"">FAIL</span>: " & testName & "<br>"
		end if
	end sub

	dim jsonLiteral
	jsonLiteral = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"

	' ============================================================
	' Test: parse a JSON object
	' ============================================================
	%><h3>Test: parse a JSON object</h3><%
	dim parseObj, parseStart
	set parseObj = new JSONobject
	parseObj.debug = false

	parseStart = timer()
	parseObj.parse jsonLiteral
	response.write "Parse time: " & round(timer() - parseStart, 6) & " s<br>"

	Call checkResult(parseObj.value("strings") = "valorTexto", "parse object: string value")
	Call checkResult(parseObj.value("numbers") = 123.456, "parse object: number value")
	Call checkResult(parseObj.value("bools") = true, "parse object: boolean value")
	Call checkResult(InStr(parseObj.Serialize(), """arrays"":[1,""2"",3.4") > 0, "parse object: array value")
	Call checkResult(InStr(parseObj.Serialize(), """emptyArray"":[]") > 0, "parse object: empty array")
	Call checkResult(InStr(parseObj.Serialize(), """emptyObject"":{}") > 0, "parse object: empty object")
	Call checkResult(InStr(parseObj.Serialize(), """prop1"":""outroTexto""") > 0, "parse object: nested object")
	Call checkResult(InStr(parseObj.Serialize(), "Texto com") > 0, "parse object: multiline value")

	' ============================================================
	' Test: parse a JSON array
	' ============================================================
	%><h3>Test: parse a JSON array</h3><%
	dim arrayLiteral, arrayParser, parsedArray
	arrayLiteral = "[" & jsonLiteral & "]"

	set arrayParser = new JSONobject
	set parsedArray = arrayParser.parse(arrayLiteral)

	Call checkResult(TypeName(parsedArray) = "JSONarray", "parse array: returns a JSONarray")
	Call checkResult(parsedArray.length = 1, "parse array: has one item")
	Call checkResult(InStr(parsedArray.Serialize(), "valorTexto") > 0, "parse array: item serialized")

	' ============================================================
	' Test: add properties
	' ============================================================
	%><h3>Test: add properties</h3><%
	dim addObj, simpleArr, multArr, nested
	set addObj = new JSONobject
	addObj.debug = false

	addObj.add "nome", "José"
	addObj.add "ficticio", true
	addObj.add "idade", 25
	addObj.add "saldo", -52
	addObj.add "bio", "Nascido em São Paulo\Brasil" & vbcrlf & "Sem filhos" & vbcrlf & vbtab & "Jogador de WoW"
	addObj.add "data-atual", now

	simpleArr = Array(1, "teste", 234.56, "mais teste", "234", now)
	addObj.add "lista", simpleArr

	redim multArr(2, 3)
	multArr(0, 0) = "0,0"
	multArr(0, 1) = "0,1"
	multArr(0, 2) = "0,2"
	multArr(0, 3) = "0,3"
	multArr(1, 0) = "1,0"
	multArr(1, 1) = "1,1"
	multArr(1, 2) = "1,2"
	multArr(1, 3) = "1,3"
	multArr(2, 0) = "2,0"
	multArr(2, 1) = "2,1"
	multArr(2, 2) = "2,2"
	multArr(2, 3) = "2,3"
	addObj.add "lista2", multArr

	set nested = new JSONobject
	nested.add "sub1", "value of sub1"
	nested.add "sub2", "value of ""sub2"""
	addObj.add "nested", nested

	Call checkResult(addObj.value("nome") = "José", "add: string property")
	Call checkResult(addObj.value("ficticio") = true, "add: boolean property")
	Call checkResult(addObj.value("idade") = 25, "add: integer property")
	Call checkResult(addObj.value("saldo") = -52, "add: negative integer property")
	Call checkResult(InStr(addObj.Serialize(), "Nascido em São Paulo") > 0, "add: multiline string property")
	Call checkResult(InStr(addObj.Serialize(), """lista"":[1,""teste"",234.56") > 0, "add: simple array property")
	Call checkResult(InStr(addObj.Serialize(), """lista2"":[[""0,0""") > 0, "add: multidimensional array property")
	Call checkResult(InStr(addObj.Serialize(), """sub1"":""value of sub1""") > 0, "add: nested object property")
	Call checkResult(InStr(addObj.Serialize(), """sub2"":") > 0, "add: escaped quotes in nested object")

	' ============================================================
	' Test: remove properties
	' ============================================================
	%><h3>Test: remove properties</h3><%
	dim removeObj
	set removeObj = new JSONobject
	removeObj.add "removeMe", "x"
	removeObj.add "keepMe", "y"
	removeObj.remove "removeMe"
	removeObj.remove "doesNotExist"

	Call checkResult(InStr(removeObj.Serialize(), "removeMe") = 0, "remove: existing property removed")
	Call checkResult(removeObj.value("keepMe") = "y", "remove: other property untouched")
	Call checkResult(IsNull(removeObj.value("doesNotExist")), "remove: non-existent property is silent")

	' ============================================================
	' Test: get values
	' ============================================================
	%><h3>Test: get values</h3><%
	dim getObj
	set getObj = new JSONobject
	getObj.add "nome", "José"
	getObj.add "idade", 25

	Call checkResult(getObj.value("nome") = "José", "value: value() method")
	Call checkResult(getObj("idade") = 25, "value: default property syntax")
	Call checkResult(IsNull(getObj("aNonExistantPropertyName")), "value: non-existent returns Null")

	' ============================================================
	' Test: change values
	' ============================================================
	%><h3>Test: change values</h3><%
	dim changeObj
	set changeObj = new JSONobject
	changeObj.add "nome", "José"
	changeObj.change "nome", "Mario"
	changeObj.change "created", -1

	Call checkResult(changeObj.value("nome") = "Mario", "change: existing property")
	Call checkResult(changeObj.value("created") = -1, "change: creates missing property")

	' ============================================================
	' Test: array push
	' ============================================================
	%><h3>Test: array push</h3><%
	dim pushArr, pushedObj, firstItem
	set pushArr = new JSONarray
	set pushedObj = new JSONobject
	pushedObj.add "newJson", "property"
	pushedObj.add "version", pushedObj.version

	pushArr.Push pushedObj
	pushArr.Push 1
	pushArr.Push "strings too"

	set firstItem = pushArr(0)

	Call checkResult(pushArr.length = 3, "push: length after three pushes")
	Call checkResult(firstItem.value("newJson") = "property", "push: object item")
	Call checkResult(pushArr(1) = 1, "push: integer item")
	Call checkResult(pushArr(2) = "strings too", "push: string item")

	' ============================================================
	' Test: load recordsets
	' ============================================================
	%><h3>Test: load recordsets</h3>
	<!--
	   METADATA
	   TYPE="TypeLib"
	   NAME="Microsoft ActiveX Data Objects 2.5 Library"
	   UUID="{00000205-0000-0010-8000-00AA006D2EA4}"
	   VERSION="2.5"
	-->
	<%
	dim rs, firstObj, rsArr
	set rs = createObject("ADODB.Recordset")
	rs.CursorType = adOpenKeyset
	rs.CursorLocation = adUseClient
	rs.LockType = adLockOptimistic

	rs.Fields.Append "ID", adInteger, , adFldKeyColumn
	rs.Fields.Append "NomeCompleto", adVarChar, 50, adFldMayBeNull
	rs.Fields.Append "Valor", adDecimal, 14, adFldMayBeNull
	rs.Fields("Valor").NumericScale = 2

	rs.Open

	rs.AddNew
	rs("ID") = 1
	rs("NomeCompleto") = "Nome 1"
	rs("Valor") = 10.99
	rs.Update

	rs.AddNew
	rs("ID") = 2
	rs("NomeCompleto") = "Nome 2"
	rs("Valor") = 29.90
	rs.Update

	set firstObj = new JSONobject
	rs.moveFirst
	firstObj.LoadFirstRecord rs

	Call checkResult(firstObj.value("ID") = 1, "LoadFirstRecord: integer field")
	Call checkResult(firstObj.value("NomeCompleto") = "Nome 1", "LoadFirstRecord: string field")
	Call checkResult(Abs(CDbl(firstObj.value("Valor")) - 10.99) < 0.001, "LoadFirstRecord: decimal field")

	set rsArr = new JSONarray
	rs.moveFirst
	rsArr.LoadRecordSet rs

	dim rsFirst
	set rsFirst = rsArr(0)

	Call checkResult(rsArr.length = 2, "JSONarray.LoadRecordSet: row count")
	Call checkResult(rsFirst.value("NomeCompleto") = "Nome 1", "JSONarray.LoadRecordSet: first row")

	set firstObj = new JSONobject
	rs.moveFirst
	firstObj.LoadRecordSet rs

	Call checkResult(InStr(firstObj.Serialize(), """NomeCompleto"":""Nome 1""") > 0, "JSONobject.LoadRecordSet: first row")
	Call checkResult(InStr(firstObj.Serialize(), """NomeCompleto"":""Nome 2""") > 0, "JSONobject.LoadRecordSet: second row")

	rs.Close
	set rs = nothing

	' ============================================================
	' Test: default property name
	' ============================================================
	%><h3>Test: default property name</h3><%
	dim defObj, defArr
	set defObj = new JSONobject
	Call checkResult(defObj.defaultPropertyName = "data", "defaultPropertyName: initial value")
	defObj.defaultPropertyName = "CustomName"
	Call checkResult(defObj.defaultPropertyName = "CustomName", "defaultPropertyName: setter")

	set defArr = new JSONarray
	defArr.defaultPropertyName = "CustomArrName"
	Call checkResult(defArr.defaultPropertyName = "CustomArrName", "JSONarray.defaultPropertyName: setter")

	' ============================================================
	' Test: lowerCaseKeys
	' ============================================================
	%><h3>Test: lowerCaseKeys</h3><%
	dim lcObj, lcArr, lcInner, lcJson
	set lcObj = new JSONobject
	lcObj.add "FULLNAME", "John"
	lcObj.lowerCaseKeys = true
	lcJson = lcObj.Serialize()

	Call checkResult(InStr(lcJson, """fullname""") > 0, "lowerCaseKeys: JSONobject key lowercased")
	Call checkResult(lcObj.value("FULLNAME") = "John", "lowerCaseKeys: original lookup still works")

	set lcArr = new JSONarray
	set lcInner = new JSONobject
	lcInner.add "ID", 123
	lcArr.Push lcInner
	lcArr.lowerCaseKeys = true
	lcJson = lcArr.Serialize()

	Call checkResult(InStr(lcJson, """id""") > 0, "lowerCaseKeys: JSONarray item key lowercased")

	set lcObj = new JSONobject
	lcObj.add JSON_ROOT_KEY, "v"
	lcObj.defaultPropertyName = "RootKey"
	lcObj.lowerCaseKeys = true
	lcJson = lcObj.Serialize()

	Call checkResult(InStr(lcJson, """rootkey""") > 0, "lowerCaseKeys: wrapper key lowercased")

	' ============================================================
	' Test: JSONscript serialization
	' ============================================================
	%><h3>Test: JSONscript serialization</h3><%
	dim jsCode, jsKey, jsObj, jsScr, expectedJs, actualJs
	jsCode = "function() { alert('test'); }"
	jsKey = "script"
	expectedJs = "{""" & jsKey & """:" & jsCode & "}"

	set jsScr = new JSONscript
	jsScr.value = jsCode
	set jsObj = new JSONobject
	jsObj.Add jsKey, jsScr
	actualJs = jsObj.Serialize()

	Call checkResult(actualJs = expectedJs, "JSONscript: serialized without quotes")

	' ============================================================
	' Test: Issue regressions
	' ============================================================
	%><h3>Test: Issue regressions</h3><%

	' Issue #103: 2D VBScript arrays serialize with inter-row commas
	dim arr2d(1, 1), issue103Obj, issue103Json
	arr2d(0, 0) = 1
	arr2d(0, 1) = 2
	arr2d(1, 0) = 3
	arr2d(1, 1) = 4
	set issue103Obj = new JSONobject
	issue103Obj.add "matrix", arr2d
	issue103Json = issue103Obj.Serialize()
	Call checkResult(InStr(issue103Json, "[[1,2],[3,4]]") > 0, "Issue #103: 2D array serialization has inter-row commas")

	' Issue #104: Parse() function call syntax - verify correct usage works
	dim issue104Obj, issue104Result
	set issue104Obj = new JSONobject
	set issue104Result = issue104Obj.parse("{""a"":1}")
	Call checkResult(issue104Result.value("a") = 1, "Issue #104: Parse() with parentheses works")

	' Issue #94: Can't access array after parse (root-level array)
	dim issue94Parser, issue94Arr
	set issue94Parser = new JSONobject
	set issue94Arr = issue94Parser.parse("[{""id"":1},{""id"":2}]")
	Call checkResult(TypeName(issue94Arr) = "JSONarray", "Issue #94: Root array parses as JSONarray")
	Call checkResult(issue94Arr.length = 2, "Issue #94: Root array has correct length")
	Call checkResult(issue94Arr(0).value("id") = 1, "Issue #94: Root array item access")

	' Issue #68: Array not closed correctly in parse (nested arrays)
	dim issue68Parser, issue68Obj
	set issue68Parser = new JSONobject
	set issue68Obj = issue68Parser.parse("{""a"":[[1,2],[3,4]]}")
	Call checkResult(InStr(issue68Obj.Serialize(), """a"":[[1,2],[3,4]]") > 0, "Issue #68: Nested arrays parse correctly")

	' Issue #66: Date serialization produces valid JSON
	dim issue66Obj, issue66Json
	set issue66Obj = new JSONobject
	issue66Obj.add "dt", #1/2/2020 3:04:05 AM#
	issue66Json = issue66Obj.Serialize()
	Call checkResult(InStr(issue66Json, "2020-01-02T03:04:05") > 0, "Issue #66: Date serializes to ISO format")
	Call checkResult(Right(issue66Json, 2) = """" & "}", "Issue #66: Date JSON is well-formed (ends with quoted string)")

	' Issue #45: LCID handling with arrays
	dim issue45Obj, issue45Json, oldLCID
	oldLCID = Response.LCID
	Response.LCID = 1033 ' en-US
	set issue45Obj = new JSONobject
	issue45Obj.add "num", 1234.56
	issue45Json = issue45Obj.Serialize()
	Response.LCID = oldLCID
	Call checkResult(InStr(issue45Json, "1234.56") > 0, "Issue #45: Numbers use dot decimal separator with en-US LCID")

	' ============================================================
	' Summary
	' ============================================================
	%><h3>Summary</h3><%
	response.write "Passed: " & passCount & "<br>"
	response.write "Failed: " & failCount & "<br>"
	if failCount = 0 then
		response.write "<b>All " & passCount & " assertions passed.</b><br>"
	else
		response.write "<b>" & failCount & " assertion(s) failed.</b><br>"
	end if
	%>
</body>
</html>
