<%
Option Explicit
Response.LCID = 1033
Response.Buffer = true
Server.ScriptTimeout = 300
%>
<!--#include file="jsonObject.class.asp"-->
<% Response.ContentType = "text/plain" %>
<%
dim passCount, failCount
passCount = 0
failCount = 0

sub checkResult(cond, testName)
	if cond then
		passCount = passCount + 1
		Response.Write "PASS: " & testName & vbCrLf
	else
		failCount = failCount + 1
		Response.Write "FAIL: " & testName & vbCrLf
	end if
	Response.Flush
end sub

sub runFixture(fileName, rootKey, expectedCount, expectedMetaCount, doParse)
	dim fso, f, path, s, parser, obj, arr, meta, firstItem, tags, nested
	dim parseStart, parseSeconds, parseErr

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

	Response.Write "Reading " & fileName & "..." & vbCrLf
	Response.Flush

	set f = fso.OpenTextFile(path, 1)
	s = f.ReadAll()
	f.Close
	set f = nothing

	Call checkResult(Len(s) > 0, fileName & ": content read")

	Response.Write "Parsing " & fileName & "..." & vbCrLf
	Response.Flush

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
	Response.Write "      parse time: " & FormatNumber(parseSeconds, 3) & " s" & vbCrLf
	Response.Flush

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

Response.Write "ASPJSON fixture tests" & vbCrLf & vbCrLf
Response.Flush

Call runFixture("small-file.json", "dataset", 200, 200, true)
Call runFixture("medium-file.json", "data", 10000, 10000, true)

' The large fixture is not parsed: the parser runs at roughly 0.14 MB/s,
' so a 222 MB document would take around 25 minutes. Verify existence and size only.
Call runFixture("large-file.json", "records", 150000, 150000, false)

Response.Write vbCrLf
Response.Write "Passed: " & passCount & vbCrLf
Response.Write "Failed: " & failCount & vbCrLf

if failCount = 0 then
	Response.Write "All " & passCount & " assertions passed." & vbCrLf
else
	Response.Write failCount & " assertion(s) failed." & vbCrLf
end if
%>
