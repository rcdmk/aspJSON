# JSONobject API

Represents a JSON object: an ordered collection of name/value pairs. It can parse JSON text, build objects programmatically, serialize them, and load ADO recordsets.

```asp
Dim json
Set json = New JSONobject
```

**Version:** `3.9.0`

## Members at a glance

### Properties

| Property              | Access  | Type                        | Default   | Description                                                                       |
| --------------------- | ------- | --------------------------- | --------- | --------------------------------------------------------------------------------- |
| `debug`               | get/set | `Boolean`                   | `False`   | Writes parser and serializer diagnostics to the response.                         |
| `defaultPropertyName` | get/set | `String`                    | `"data"`  | Key used when an array is stored at the root, and by `LoadRecordSet`.             |
| `depth`               | get     | `Integer`                   | `0`       | Depth in the object chain. Root is `0`; setting `parent` sets `parent.depth + 1`. |
| `pairs`               | get     | `Variant()` of `JSONpair`   | —         | Snapshot of the object's name/value pairs.                                        |
| `parent`              | get/set | `JSONobject` or `JSONarray` | `Nothing` | Parent container. Assigning it recalculates `depth`.                              |
| `version`             | get     | `String`                    | —         | Class version (`"3.9.0"`).                                                        |

### Methods

| Method                                                                                  | Returns                    | Description                                                                         |
| --------------------------------------------------------------------------------------- | -------------------------- | ----------------------------------------------------------------------------------- |
| [`Parse(strJson)`](#parse)                                                              | `JSONobject` / `JSONarray` | Parses JSON text and populates the object.                                          |
| [`Add(prop, value)`](#add)                                                              | —                          | Adds a property. Raises an error if the name exists.                                |
| [`Remove(prop)`](#remove)                                                               | —                          | Removes a property; silent if it does not exist.                                    |
| [`Value(prop)`](#value)                                                                 | `Variant`                  | Returns a property value. This is the default function, so `json(prop)` also works. |
| [`Change(prop, value)`](#change)                                                        | —                          | Updates a property, creating it if necessary.                                       |
| [`Serialize()`](#serialize)                                                             | `String`                   | Returns the JSON representation.                                                    |
| [`Write()`](#write)                                                                     | —                          | Writes the JSON representation to the response.                                     |
| [`LoadRecordSet(rs)`](#loadrecordset)                                                   | —                          | Loads all records into an array under `defaultPropertyName`.                        |
| [`LoadFirstRecord(rs)`](#loadfirstrecord)                                               | —                          | Adds the fields of the current record as properties.                                |
| [`ArrayPush(arr, value)`](#arraypush-deprecated)                                        | `Variant()`                | **Deprecated.** Appends a value to a VBScript array.                                |
| [`GetTypeName(value)`](#gettypename)                                                    | `String`                   | Type name with extra support for MySQL value types.                                 |
| [`EscapeCharacters(text)`](#escapecharacters)                                           | `String`                   | Escapes a string for JSON output.                                                   |
| [`SerializeObject(obj)`](#serializeobject-serializevalue-serializearray-numdimensions)  | `String`                   | Serializes a `JSONobject`.                                                          |
| [`SerializeValue(value)`](#serializeobject-serializevalue-serializearray-numdimensions) | `String`                   | Serializes a scalar value.                                                          |
| [`SerializeArray(arr)`](#serializeobject-serializevalue-serializearray-numdimensions)   | `String`                   | Serializes a `JSONarray` or VBScript array.                                         |
| [`NumDimensions(arr)`](#serializeobject-serializevalue-serializearray-numdimensions)    | `Integer`                  | Returns the number of dimensions of a VBScript array.                               |

## Properties

### debug

```asp
json.debug = True   ' enable
json.debug = False  ' disable (default)
```

When `True`, the parser and serializer emit `<li>` log lines to the response. The output is no longer valid JSON, so use this only for troubleshooting.

### defaultPropertyName

```asp
json.defaultPropertyName = "data"   ' default
```

The key used when:

- `Parse` encounters an array at the JSON root (the parsed array is stored under this name on the object).
- `LoadRecordSet` stores the record array on the object.
- A `JSONobject` serializes a root array.

Property lookup has a special fallback: `json(defaultPropertyName)` returns the root value even though the internal pair uses a sentinel name. A real property with the same name takes precedence.

```asp
Set json = New JSONobject
json.Parse "[1,2,3]"

json.Write()                        ' {"data":[1,2,3]}
Response.Write json("data").length  ' 3

json.defaultPropertyName = "items"
json.Write()                        ' {"items":[1,2,3]}
```

### depth

Read-only. The root object has depth `0`. A nested object or array gets depth `parent.depth + 1` when its `parent` property is assigned.

```asp
Set child = New JSONobject
Set child.parent = json
Response.Write child.depth          ' 1
```

### pairs

Returns a snapshot array of `JSONpair` objects, trimmed to the current property count. It is intended for iteration and inspection; modifying the returned array does not modify the object.

```asp
Dim pair
For Each pair In json.pairs
    Response.Write pair.name
Next
```

### parent

The parent `JSONobject` or `JSONarray`, or `Nothing` for a root object.

```asp
Set json.parent = parentObject
Response.Write json.depth
```

Assigning a parent recalculates `depth`. The class uses this relationship during parsing and while resolving the root array.

### version

```asp
Response.Write json.version          ' 3.9.0
```

## Methods

### Parse

Parses a JSON-formatted string.

```asp
Set result = json.Parse(strJson)
```

**Parameters**

| Name      | Type     | Description         |
| --------- | -------- | ------------------- |
| `strJson` | `String` | JSON text to parse. |

**Returns:** the root of the parsed document.

- If the root is an object, `json` itself is cleared and repopulated, and the method returns `json`.
- If the root is an array, a new `JSONarray` is returned. The array is also stored on `json` under `defaultPropertyName`, so `json.Write()` produces `{"data":[...]}`.

```asp
Dim json, parsed, rows
Set json = New JSONobject

' Object root
Set parsed = json.Parse("{""a"":1,""b"":[true,false]}")
Response.Write parsed("a")          ' 1

' Array root
Set rows = json.Parse("[{""id"":1},{""id"":2}]")
Response.Write rows.length          ' 2
json.Write()                        ' {"data":[{"id":1},{"id":2}]}
```

**Remarks**

- The input is trimmed before parsing.
- Strings shorter than two characters raise `JSON_ERROR_PARSE` (`1`).
- Objects and arrays become `JSONobject` and `JSONarray` instances.
- Numbers become `Double`, `true`/`false` become `Boolean`, `null` becomes `Null`, and `undefined` is evaluated with `Eval` (normally becoming `Empty`).
- `\uXXXX` escapes and common string escapes are decoded. See [serialization → parsing](../reference/serialization.md#parsing-type-mapping).
- Duplicate keys raise `JSON_ERROR_PROPERTY_ALREADY_EXISTS` (`2`).
- The method temporarily sets `Response.LCID` to `1033` and restores the previous value.
- A bare scalar root (for example `"123"`) is not supported; use an object or array root.

### Add

Adds a new name/value pair.

```asp
json.Add prop, value
```

**Parameters**

| Name    | Type      | Description     |
| ------- | --------- | --------------- |
| `prop`  | `String`  | Property name.  |
| `value` | `Variant` | Value to store. |

**Remarks**

- Raises `JSON_ERROR_PROPERTY_ALREADY_EXISTS` (`2`) if a property with the same name exists.
- A VBScript array is converted to a `JSONarray`.
- An ADO `Field` is unwrapped to its `.Value`.
- Other objects are stored by reference; scalars are stored by value.
- `IStringList` values (for example some request collections) are stored as scalars instead of being treated as nested objects.

```asp
json.Add "name", "Ada"
json.Add "tags", Array("math", "code")

Set nested = New JSONobject
nested.Add "city", "Sao Paulo"
json.Add "address", nested
```

### Remove

Removes a property.

```asp
json.Remove prop
```

If the property does not exist, the call does nothing and does not raise an error. Removing a property shifts the remaining properties down.

```asp
json.Remove "name"
json.Remove "doesNotExist"   ' no error
```

### Value

Returns the value of a property. `Value` is the class's default function, so `json("name")` is equivalent to `json.Value("name")`.

```asp
value = json.Value(prop)
value = json(prop)
```

**Returns:** the property value, or `Null` if it does not exist. Object values must be assigned with `Set`:

```asp
Response.Write json("name")

Dim child
Set child = json("address")
child.Write()
```

If `prop` equals `defaultPropertyName` and no real property has that name, `Value` returns the root value (the parsed root array or the recordset array).

### Change

Updates a property, creating it when it does not exist.

```asp
json.Change prop, value
```

**Parameters**

| Name    | Type      | Description    |
| ------- | --------- | -------------- |
| `prop`  | `String`  | Property name. |
| `value` | `Variant` | New value.     |

**Remarks**

- If the property exists, its value is replaced.
- If it does not exist, it is added (equivalent to `Add`).
- VBScript arrays are converted to `JSONarray`; other objects are assigned by reference.

```asp
json.Change "name", "Grace"
json.Change "newProp", 123
```

### Serialize

Returns the JSON representation of the object.

```asp
jsonString = json.Serialize()
```

The output is compact (no indentation or extra whitespace). Serialization:

- Temporarily sets `Response.LCID` to `1033` so numbers use `.` as the decimal separator, then restores it.
- Escapes strings as described in [serialization → escaping](../reference/serialization.md#string-escaping).
- Serializes `Date` values as `"YYYY-MM-DDTHH:MM:SS"`.
- Recursively serializes nested objects and arrays.

```asp
json.Serialize()
' {"name":"Ada","tags":["math","code"]}
```

### Write

Writes `Serialize()` directly to the response.

```asp
json.Write()
' equivalent to: Response.Write json.Serialize()
```

### LoadRecordSet

Loads every remaining record of an ADO recordset into an array stored under `defaultPropertyName`.

```asp
json.LoadRecordSet rs
```

**Parameters**

| Name | Type              | Description     |
| ---- | ----------------- | --------------- |
| `rs` | `ADODB.Recordset` | Open recordset. |

**Remarks**

- Reads from the current position to EOF and leaves the recordset at EOF.
- Each record becomes a `JSONobject` whose properties match the field names.
- The resulting `JSONarray` is assigned with `Change`, replacing an existing value under `defaultPropertyName`.
- Database `Null` values are preserved as `Null`; date fields become `Date` values.
- Call `rs.MoveFirst` before loading if the recordset may not be positioned at the first record.

```asp
Set json = New JSONobject
json.LoadRecordSet rs
json.Write()
' {"data":[{"id":1,"name":"Ada"},{"id":2,"name":"Grace"}]}
```

> For a bare array instead of `{"data":[...]}`, use [`JSONarray.LoadRecordSet`](jsonarray.md#loadrecordset).

### LoadFirstRecord

Adds the fields of the **current** record as properties of the object.

```asp
json.LoadFirstRecord rs
```

**Remarks**

- Does not move the recordset cursor.
- Uses `Add`, so an existing property with the same name raises `JSON_ERROR_PROPERTY_ALREADY_EXISTS` (`2`).
- Useful for returning a single row as a flat object.

```asp
rs.MoveFirst
json.LoadFirstRecord rs
json.Write()
' {"id":1,"name":"Ada"}
```

### ArrayPush (deprecated)

```asp
arr = json.ArrayPush(arr, value)
```

Appends a value to a VBScript array and returns the resized array. **Deprecated** and scheduled for removal; use [`JSONarray.Push`](jsonarray.md#push) instead.

```asp
Dim arr
arr = Array()
arr = json.ArrayPush(arr, "first")
arr = json.ArrayPush(arr, "second")
```

### GetTypeName

Returns the VBScript type name of a value, with fallbacks for ADO/MySQL types that `TypeName` does not report correctly.

```asp
typeName = json.GetTypeName(value)
```

| VarType | Returned name               |
| ------- | --------------------------- |
| `14`    | `"Decimal"` (MySQL decimal) |
| `16`    | `"Integer"` (MySQL tinyint) |

All other values return `TypeName(value)`.

### EscapeCharacters

Escapes a string for inclusion in JSON output.

```asp
escaped = json.EscapeCharacters(text)
```

Escapes `\`, `"`, CR, LF, TAB, and backspace. It does not escape `/` or every control character. A `Null` input is returned unchanged.

```asp
json.EscapeCharacters("line1" & vbCrLf & "line2")
' line1\nline2
```

### SerializeObject, SerializeValue, SerializeArray, NumDimensions

These helpers are public but are primarily implementation-level building blocks. They may change between versions.

```asp
json.SerializeObject(obj)      ' serializes a JSONobject
json.SerializeValue(value)     ' serializes a scalar
json.SerializeArray(arr)       ' serializes a JSONarray or a 1D/2D VBScript array
json.NumDimensions(arr)        ' number of dimensions of a VBScript array
```

`SerializeValue(Nothing)` returns `undefined`, but normal object/array serialization of `Nothing` raises `424` (`Object required`). Use `Null` for JSON `null`.

`SerializeArray` handles one- and two-dimensional arrays. Two-dimensional arrays are emitted as an array of arrays. Three or more dimensions are not supported: `Add` accepts them, but serialization raises error `9` (`Subscript out of range`), and cleanup may raise the same error.

## Example

```asp
<%@ Language="VBScript" LCID="1033" %>
<% Option Explicit %>
<!--#include file="jsonObject.class.asp" -->
<%
Dim json, address, tags
Set json = New JSONobject

Set address = New JSONobject
address.Add "city", "Sao Paulo"

Set tags = New JSONarray
tags.Push "json"
tags.Push "asp"

json.Add "name", "Ada"
json.Add "active", True
json.Add "address", address
json.Add "tags", tags
json.Add "created", Now

Response.ContentType = "application/json"
json.Write()
%>
```

## Related

- [JSONarray API](jsonarray.md)
- [Serialization reference](../reference/serialization.md)
- [Error reference](../reference/errors.md)
