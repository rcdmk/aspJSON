# Usage guide

This guide covers the common workflows. For exhaustive member lists, see the [JSONobject API](api/jsonobject.md) and [JSONarray API](api/jsonarray.md).

## Create an object and add values

```asp
Dim json
Set json = New JSONobject

json.Add "name", "Ada"
json.Add "age", 36
json.Add "active", True
json.Add "score", 99.5
json.Add "missing", Null

json.Write()
' {"name":"Ada","age":36,"active":true,"score":99.5,"missing":null}
```

`Add` accepts strings, numbers, booleans, `Null`, dates, arrays, `JSONobject`, `JSONarray`, `JSONscript`, and ADO `Field` values.

### Nested objects

```asp
Dim address
Set address = New JSONobject
address.Add "city", "Sao Paulo"
address.Add "country", "Brazil"

json.Add "address", address
```

### Arrays

A VBScript array is converted to a `JSONarray` automatically:

```asp
json.Add "numbers", Array(1, 2, 3)
json.Add "mixed", Array("one", 2, True, Null)
```

You can also build a `JSONarray` directly when you need to grow it dynamically:

```asp
Dim tags
Set tags = New JSONarray
tags.Push "json"
tags.Push "asp"
tags.Push 3.9

json.Add "tags", tags
```

### Multidimensional arrays

Two-dimensional VBScript arrays are serialized as an array of arrays:

```asp
Dim grid
grid = Array(Array(1, 2), Array(3, 4))
json.Add "grid", grid
' "grid":[[1,2],[3,4]]
```

Only one- and two-dimensional arrays are handled by the serializer.

## Read values

Use `Value(prop)` or the default-property shorthand `json(prop)`:

```asp
Response.Write json.Value("name")     ' Ada
Response.Write json("age")            ' 36
```

A missing property returns `Null` instead of raising an error:

```asp
If IsNull(json("doesNotExist")) Then
    Response.Write "not found"
End If
```

When a property contains an object, assign it with `Set`:

```asp
Dim address
Set address = json("address")
Response.Write address("city")

' Or serialize it directly:
json("address").Write()
```

### Iterate properties

The `pairs` property returns a snapshot array of `JSONpair` objects:

```asp
Dim pair, value
For Each pair In json.pairs
    If IsObject(pair.value) Then
        Set value = pair.value
        Response.Write pair.name & ": " & value.Serialize()
    Else
        Response.Write pair.name & ": " & pair.value
    End If

    Response.Write "<br>"
Next
```

> Property names are compared with VBScript's default binary comparison, so lookups are case-sensitive unless the page changes `Option Compare`.

## Change and remove values

`Change` updates an existing property and **creates it if it does not exist**:

```asp
json.Change "name", "Grace"     ' update
json.Change "newProp", 123      ' create
```

`Remove` deletes a property. Removing a missing property is a silent no-op:

```asp
json.Remove "newProp"
json.Remove "neverExisted"      ' no error
```

## The default property name

`defaultPropertyName` controls the key used when an array is stored at the root of a `JSONobject`. It defaults to `"data"`.

```asp
Dim json
Set json = New JSONobject
json.Parse "[1,2,3]"

json.Write()
' {"data":[1,2,3]}

json.defaultPropertyName = "items"
json.Write()
' {"items":[1,2,3]}

json("items").Write()
' [1,2,3]
```

The same property is used by `LoadRecordSet`:

```asp
json.defaultPropertyName = "rows"
json.LoadRecordSet rs
json.Write()
' {"rows":[...]}
```

When the JSON root is an object, it is loaded directly into the object and no wrapper property is added.

## Work with arrays

`JSONarray` has a default function `ItemAt`, so index access uses parentheses:

```asp
Dim arr
Set arr = New JSONarray

arr.Push "first"
arr.Push 2
arr.Push True

Response.Write arr(0)          ' first
Response.Write arr.length      ' 3
Response.Write arr.capacity    ' internal buffer size, >= length

' Object items must be assigned with Set:
Dim nested
Set nested = arr(2)
```

Accessing an index outside `0 .. length - 1` raises error `9` (`JSON_ERROR_INDEX_OUT_OF_BOUNDS`).

### Loop over an array

Readable version, using the `items` snapshot:

```asp
Dim item
For Each item In arr.items
    If IsObject(item) Then
        item.Write()
    Else
        Response.Write item
    End If

    Response.Write "<br>"
Next
```

Faster version, using index access:

```asp
Dim i, item
For i = 0 To arr.length - 1
    If IsObject(arr(i)) Then
        Set item = arr(i)
        item.Write()
    Else
        Response.Write arr(i)
    End If

    Response.Write "<br>"
Next
```

`items` returns a trimmed copy of the internal buffer, so changing it does not modify the array. Use `Push` to add items.

## Parse JSON text

```asp
Dim json, parsed
Set json = New JSONobject
Set parsed = json.Parse("{""name"":""Ada"",""tags"":[""math"",""code""]}")

Response.Write parsed("name")            ' Ada

Dim tags
Set tags = parsed("tags")
Response.Write tags(0)                  ' math
Response.Write tags.length              ' 2
```

If the root is an array, `Parse` returns a `JSONarray`:

```asp
Dim json, rows
Set json = New JSONobject
Set rows = json.Parse("[{""id"":1},{""id"":2}]")

Response.Write rows.length              ' 2
rows.Write()                            ' [{"id":1},{"id":2}]

' The same data is also available as the default property:
json.Write()                            ' {"data":[{"id":1},{"id":2}]}
```

If the root is an object, the current object is repopulated and returned:

```asp
json.Parse "{""a"":1}"
json.Write()                            ' {"a":1}
```

Notes:

- Leading and trailing whitespace is trimmed.
- Object and array roots are supported. A bare scalar root is not parsed into a value.
- Duplicate keys raise error `2` (`JSON_ERROR_PROPERTY_ALREADY_EXISTS`).
- Strings and numbers are converted as described in the [serialization reference](reference/serialization.md#parsing-type-mapping).
- A JSON string shorter than two characters raises error `1` (`JSON_ERROR_PARSE`).

## Load a recordset

### As an object with a data property

```asp
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
' {"data":[{"id":1,"name":"..."},{"id":2,"name":"..."}]}
```

`LoadRecordSet` reads from the **current position to EOF** and leaves the recordset at EOF. Call `rs.MoveFirst` if you need to read it again.

### As a bare array

```asp
Set rs = cn.Execute("SELECT id, name FROM customers ORDER BY id")

Dim rows
Set rows = New JSONarray
rows.LoadRecordSet rs

rows.Write()
' [{"id":1,"name":"..."},{"id":2,"name":"..."}]
```

`JSONarray.LoadRecordSet` appends records to the existing array.

### The current record only

```asp
rs.MoveFirst

Dim json
Set json = New JSONobject
json.LoadFirstRecord rs

json.Write()
' {"id":1,"name":"..."}
```

`LoadFirstRecord` adds the fields of the current row as properties. It does not move the cursor. Calling it twice on the same object raises error `2` because the properties already exist.

### Multiple recordsets on one object

Change `defaultPropertyName` before each load to keep the results in separate properties:

```asp
Set json = New JSONobject

json.defaultPropertyName = "customers"
json.LoadRecordSet rsCustomers

json.defaultPropertyName = "orders"
json.LoadRecordSet rsOrders

json.Write()
' {"customers":[...],"orders":[...]}
```

### Column types

- `Null` database values become JSON `null`.
- `adDate`/`adDBTimeStamp` values become VBScript `Date` values and serialize as ISO-like date strings.
- MySQL `DECIMAL` and `TINYINT` values are detected through `GetTypeName` (see [serialization reference](reference/serialization.md#parsing-type-mapping)).

## Emit raw text with JSONscript

Use `JSONscript` when a property must contain raw text rather than a JSON string. This is useful for embedding JavaScript callbacks:

```asp
Dim script, json
Set script = New JSONscript
script.value = "function () { alert('test'); }"

Set json = New JSONobject
json.Add "callback", script

json.Write()
' {"callback":function () { alert('test'); }}
```

`JSONscript` performs no escaping or validation, so only use it with trusted content. An empty string is written as `null`.

> `JSONscript` is handled specially when it is a property value of a `JSONobject`. If placed inside a `JSONarray`, it is serialized as an ordinary quoted string.

## Dates and special values

```asp
json.Add "created", Now
' "created":"2026-11-02T14:05:09"
```

Dates are written as `YYYY-MM-DDTHH:MM:SS` using server-local time, without milliseconds or a timezone offset.

| VBScript value   | JSON output                |
| ---------------- | -------------------------- |
| `Null`           | `null`                     |
| `Nothing`        | `undefined` (non-standard) |
| `Empty`          | `""`                       |
| `True` / `False` | `true` / `false`           |
| `Now` / a `Date` | `"YYYY-MM-DDTHH:MM:SS"`    |

Use `Null` for JSON `null`. Avoid serializing `Nothing` unless the consumer accepts non-standard `undefined`.

## Debug mode

Set `debug` to `True` to have the parser and serializer write diagnostic `<li>` lines to the response:

```asp
json.debug = True
json.Parse jsonString
```

Debug output is mixed into the response and makes the result invalid JSON. Use it only while troubleshooting, and turn it off before calling `Write` in production.

## Error handling

```asp
Dim parsed
On Error Resume Next
Set parsed = json.Parse inputString
If Err.Number <> 0 Then
    Response.Write "Error " & Err.Number & ": " & Err.Description
    Err.Clear
End If
On Error GoTo 0
```

See the [error reference](reference/errors.md) for every code and its causes.

## Performance tips

- Enable output buffering (`Response.Buffer = True`) and flush periodically for large responses.
- Turn off `debug` in production.
- Prefer `JSONarray.Push` over the deprecated `JSONobject.ArrayPush`.
- Use indexed `For` loops with `arr(i)` when iterating large arrays; use `For Each arr.items` for readability.
- `capacity` is read-only and shows the allocated buffer. It can be larger than `length`; the class grows buffers by roughly 20% to reduce reallocations.
- Parsing and serialization hold the entire document in memory. For very large payloads, raise `Server.ScriptTimeout` and the ASP buffering limit (see [`web.config`](../web.config)).
- Create objects per request. Do not store a `JSONobject` in `Application` state, because instances hold parse state and references to parents.
