# Getting started

## Requirements

- **IIS with classic ASP enabled.** The library is written in VBScript and must run inside an ASP page.
- **`Response` object available.** `JSONobject.Parse`, `Serialize`, and `Write` use `Response.LCID` and `Response.Write`, so the class cannot be used from `cscript`/`wscript` or from a plain VBScript host.
- **An LCID set before use.** See [Set the LCID](#set-the-lcid).
- **ADO (optional).** Only the `LoadRecordSet` / `LoadFirstRecord` methods need ADO (`ADODB.Recordset`).

## Install the include file

Copy [`jsonObject.class.asp`](../jsonObject.class.asp) into your web application and include it from each page that uses the library:

```asp
<!--#include file="jsonObject.class.asp" -->
```

Use `file=` for a path relative to the current page, or `virtual=` for a path relative to the application root:

```asp
<!--#include virtual="/includes/jsonObject.class.asp" -->
```

The include must appear before the first use of `JSONobject`, `JSONarray`, or `JSONscript`.

> The source file is UTF-8 with a BOM and CRLF line endings. Keep that encoding when editing it. If your pages output non-ASCII text, make sure the ASP code page and response charset match, for example:
>
> ```asp
> Response.CodePage = 65001
> Response.Charset = "UTF-8"
> ```

## Set the LCID

The library requires a valid **LCID** (locale identifier) on the page before it is used. This ensures VBScript performs numeric and date conversions consistently.

You can set it in any of these ways:

```asp
' 1. Page directive, for the whole page
<%@ Language="VBScript" LCID="1033" %>
```

```asp
' 2. Session, for every page in the session
Session.LCID = 1033
```

```asp
' 3. Response, from this point on the page
Response.LCID = 1033
```

Common values:

| LCID | Locale |
| --- | --- |
| 1033 | English (United States) |
| 1046 | Portuguese (Brazil) |
| 1031 | German (Germany) |
| 1036 | French (France) |
| 3082 | Spanish (Spain) |

Use the LCID for your own locale; JSON output itself always uses `.` as the decimal separator because the library switches to `1033` internally while parsing and serializing.

## Your first page

```asp
<%@ Language="VBScript" LCID="1033" %>
<% Option Explicit %>
<!--#include file="jsonObject.class.asp" -->
<%
Dim json
Set json = New JSONobject

json.Add "message", "Hello, world"
json.Add "ok", True
json.Add "count", 3

Response.ContentType = "application/json"
Response.Write json.Serialize()
%>
```

Output:

```json
{"message":"Hello, world","ok":true,"count":3}
```

`json.Write()` is shorthand for `Response.Write json.Serialize()`. Use `Serialize()` when you need the string, and `Write()` when you want to send it directly.

## Parse some JSON

```asp
<%
Dim json, parsed, tags
Set json = New JSONobject

Set parsed = json.Parse("{""name"":""Ada"",""tags"":[""math"",""code""]}")

Response.Write parsed("name")          ' Ada

Set tags = parsed("tags")
Response.Write tags(0)                 ' math
Response.Write tags.length             ' 2
%>
```

`Parse` returns a `JSONobject` for an object root and a `JSONarray` for an array root. Objects and arrays nested inside are converted to the corresponding class automatically.

## Load a recordset

```asp
<%
Dim cn, rs, json
Set cn = CreateObject("ADODB.Connection")
cn.Open "yourConnectionStringGoesHere"

Set rs = cn.Execute("SELECT id, name FROM customers ORDER BY id")

Set json = New JSONobject
json.LoadRecordSet rs

rs.Close
cn.Close
Set rs = Nothing
Set cn = Nothing

json.Write()
' {"data":[{"id":1,"name":"..."}, ...]}
%>
```

The generated array is stored under the object's `defaultPropertyName`, which is `"data"` unless changed. See the [usage guide](guide.md#load-a-recordset) for details.

## Run the test page

[`test.asp`](../test.asp) exercises parsing, adding, changing, removing, arrays, recordsets, and raw JavaScript output.

1. Put `test.asp` and `jsonObject.class.asp` in an application served by IIS.
2. Make sure ADO is available (the test builds an in-memory recordset).
3. Browse to `test.asp`. It sets `Response.LCID = 1046`; change that line if you prefer another locale.
4. The test page turns on `Response.Buffer`, sets `Server.ScriptTimeout = 10`, and includes a [`web.config`](../web.config) with larger ASP limits.

The flags at the top of the test page (`testLoad`, `testAdd`, and so on) let you disable individual sections.

## Next steps

- [Usage guide](guide.md) — common patterns and examples.
- [JSONobject API](api/jsonobject.md) — every property and method.
- [Serialization reference](reference/serialization.md) — exactly how VBScript values become JSON.
