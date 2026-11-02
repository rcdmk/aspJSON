# aspJSON documentation

`aspJSON` is a dependency-free JSON parser and serializer for **classic ASP / VBScript**. It is implemented entirely in `jsonObject.class.asp` and exposes four classes:

| Class        | Version | Purpose                                                                                                                     |
| ------------ | ------- | --------------------------------------------------------------------------------------------------------------------------- |
| `JSONobject` | 3.9.0   | A JSON object (ordered set of name/value pairs). Parses JSON strings, builds objects, serializes, and loads ADO recordsets. |
| `JSONarray`  | 2.4.0   | A JSON array. Grows dynamically, supports index access, and serializes nested values.                                       |
| `JSONscript` | 1.0.0   | A wrapper for raw text (such as JavaScript) that must be written to the output without JSON quoting.                        |
| `JSONpair`   | —       | Internal name/value pair used by `JSONobject`. Exposed through the `pairs` property.                                        |

The class versions are independent. `JSONobject.version` matches the library release; `JSONarray` and `JSONscript` have their own internal versions (`JSONscript` does not expose a public `version` property).

## Features

- Parse JSON text into `JSONobject` / `JSONarray` instances.
- Build JSON programmatically with `Add`, `Change`, `Remove`, and `JSONarray.Push`.
- Serialize objects, arrays, VB arrays, dates, booleans, and nested structures.
- Convert ADO recordsets into JSON with `LoadRecordSet` / `LoadFirstRecord`.
- Decode `\uXXXX` escapes and common string escapes.
- No external libraries or components required (ADO is only needed for the recordset helpers).

## Requirements

- Microsoft IIS with classic ASP enabled (VBScript).
- The class uses the ASP `Response` object (`Response.LCID`, `Response.Write`), so it runs inside ASP pages, not under `cscript`/`wscript`.
- The page, session, or response **LCID must be set before using the library**. See [Getting started](getting-started.md#set-the-lcid).

## Documentation map

| Document                                                  | Contents                                                                                      |
| --------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| [Getting started](getting-started.md)                     | Requirements, installing the include file, LCID setup, first page, running the test page.     |
| [Usage guide](guide.md)                                   | Common tasks: build, read, change, remove, parse, arrays, recordsets, raw scripts, debugging. |
| [JSONobject API](api/jsonobject.md)                       | Full reference for the `JSONobject` class.                                                    |
| [JSONarray API](api/jsonarray.md)                         | Full reference for the `JSONarray` class.                                                     |
| [JSONscript and JSONpair API](api/jsonscript-jsonpair.md) | Reference for the supporting classes.                                                         |
| [Serialization reference](reference/serialization.md)     | Type mapping, escaping, dates, numbers, and parsing behavior.                                 |
| [Error reference](reference/errors.md)                    | Error constants, causes, and handling patterns.                                               |
| [Internals](internals.md)                                 | Parser state machine, storage model, performance notes, and limitations.                      |

## Quick example

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
json.Write()
' {"message":"Hello, world","ok":true,"count":3}
%>
```

## License

MIT. See [`LICENCE.txt`](../LICENCE.txt).
