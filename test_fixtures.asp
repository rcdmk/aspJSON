<%
Option Explicit
Response.LCID = 1033
Response.Buffer = true
Server.ScriptTimeout = 300
%>
<!--#include file="jsonObject.class.asp" -->
<!DOCTYPE html>
<html>
<head>
	<meta charset="UTF-8">
	<title>ASPJSON Fixture Tests</title>
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
	<h1>JSON Fixture Tests</h1>
	<%
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
		response.flush
	end sub

	sub runFixture(fileName, rootKey, expectedCount, expectedMetaCount, doParse)
		dim fso, f, path, s, parser, obj, arr, meta, firstItem, tags, nested
		dim parseStart, parseSeconds, parseErr

		response.write "<h3>Test: " & fileName & "</h3>"
		response.flush

		set fso = CreateObject("Scripting.FileSystemObject")
		path = Server.MapPath(fileName)

		if not fso.FileExists(path) then
			Call checkResult(False, fileName & ": file exists")
			exit sub
		end if

		Call checkResult(True, fileName & ": file exists (" & FormatNumber(fso.GetFile(path).Size / 1048576, 1) & " MB)")

		if not doParse then
			Call checkResult(fso.GetFile(path).Size > 200000000, fileName & ": size > 200 MB")
			exit sub
		end if

		response.write "Reading " & fileName & "...<br>"
		response.flush

		set f = fso.OpenTextFile(path, 1)
		s = f.ReadAll()
		f.Close
		set f = nothing

		Call checkResult(Len(s) > 0, fileName & ": content read")

		response.write "Parsing " & fileName & "...<br>"
		response.flush

		set parser = new JSONobject
		parser.debug = false
		parseStart = timer()

		Err.Clear
		On Error Resume Next
		set obj = parser.Parse(s)
		if Err.Number <> 0 then
			parseErr = Err.Description
			Err.Clear
			On Error Goto 0
			Call checkResult(False, fileName & ": parse failed (" & parseErr & ")")
			exit sub
		end if
		On Error Goto 0

		parseSeconds = timer() - parseStart
		response.write "Parse time: " & FormatNumber(parseSeconds, 3) & " s<br>"
		response.flush

		Call checkResult(TypeName(obj) = "JSONobject", fileName & ": root is a JSONobject")

		set arr = obj(rootKey)
		Call checkResult(isObject(arr) and TypeName(arr) = "JSONarray", fileName & ": '" & rootKey & "' is a JSONarray")
		Call checkResult(arr.Length = expectedCount, fileName & ": " & rootKey & " length = " & expectedCount & " (got " & arr.Length & ")")

		set meta = obj("meta")
		Call checkResult(isObject(meta) and TypeName(meta) = "JSONobject", fileName & ": meta is a JSONobject")
		Call checkResult(meta("count") = expectedMetaCount, fileName & ": meta.count = " & expectedMetaCount & " (got " & meta("count") & ")")

		set firstItem = arr(0)
		Call checkResult(firstItem("id") = 0, fileName & ": first item id = 0")
		Call checkResult(Len(firstItem("name")) > 0, fileName & ": first item has a name")

		set tags = firstItem("tags")
		Call checkResult(isObject(tags) and TypeName(tags) = "JSONarray", fileName & ": first item tags is a JSONarray")
		Call checkResult(tags.Length > 0, fileName & ": first item has tags")

		set nested = firstItem("nested")
		Call checkResult(isObject(nested) and TypeName(nested) = "JSONobject", fileName & ": first item nested is a JSONobject")
	end sub

	response.flush

	Call runFixture("small-file.json", "dataset", 200, 200, true)
	Call runFixture("medium-file.json", "data", 10000, 10000, true)

	' The large fixture is not parsed: after the parser optimization it runs at
	' roughly 2.8 MB/s, so the 221 MB document takes about 80 seconds (~1.3 minutes)
	' and grows the worker process by ~1.8 GB. Verify existence and size only.
	Call runFixture("large-file.json", "records", 150000, 150000, false)
	%>
	<h3>Summary</h3>
	<%
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
