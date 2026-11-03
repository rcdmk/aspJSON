# JSONarray API

Represents a JSON array. Items are appended with `Push`, read through the default `ItemAt` function or the `items` property, and serialized with `Serialize` or `Write`.

```asp
Dim arr
Set arr = New JSONarray
```

**Version:** `2.4.0`

## Members at a glance

### Properties

| Property              | Access  | Type                        | Default     | Description                                                                                          |
| --------------------- | ------- | --------------------------- | ----------- | ---------------------------------------------------------------------------------------------------- |
| `version`             | get     | `String`                    | —           | Class version (`"2.4.0"`).                                                                           |
| `items`               | get/set | `Variant()`                 | empty array | Snapshot of the items (get), or replace the contents from a VBScript array (set).                    |
| `capacity`            | get     | `Integer`                   | `0`         | Allocated buffer size; always greater than or equal to `length`.                                     |
| `length`              | get     | `Integer`                   | `0`         | Number of items in the array.                                                                        |
| `depth`               | get     | `Integer`                   | `0`         | Depth in the object chain. Set when `parent` is assigned.                                            |
| `parent`              | get/set | `JSONobject` or `JSONarray` | `Nothing`   | Parent container.                                                                                    |
| `defaultPropertyName` | get/set | `String`                    | `"data"`    | Inherited from a parent object; exposed for consistency. Array output itself is always a bare array. |

### Methods

| Method                                | Returns   | Description                                                                             |
| ------------------------------------- | --------- | --------------------------------------------------------------------------------------- |
| [`Push(value)`](#push)                | —         | Appends a value.                                                                        |
| [`ItemAt(index)`](#itemat)            | `Variant` | Returns the item at an index. This is the default function, so `arr(index)` also works. |
| [`LoadRecordSet(rs)`](#loadrecordset) | —         | Appends every remaining record as an object.                                            |
| [`Serialize()`](#serialize)           | `String`  | Returns the JSON array text.                                                            |
| [`Write()`](#write)                   | —         | Writes the JSON array to the response.                                                  |

## Properties

### version

```asp
Response.Write arr.version          ' 2.4.0
```

### items

Returns a snapshot of the array, trimmed to `length`. Modifying the returned array does not modify the `JSONarray`; use `Push` to add values.

```asp
Dim item
For Each item In arr.items
    ' ...
Next
```

Assigning an array replaces the contents and resets `length` and `capacity`:

```asp
arr.items = Array("a", "b", "c")
Response.Write arr.length           ' 3
```

If the assigned value is not an array, error `4` (`JSON_ERROR_NOT_AN_ARRAY`) is raised.

### capacity

The size of the internal buffer. It can be larger than `length` because the class grows buffers by roughly 20% to reduce reallocations. Read-only.

```asp
Response.Write arr.capacity
```

### length

The number of items. Read-only.

```asp
For i = 0 To arr.length - 1
    Response.Write arr(i)
Next
```

### depth

Read-only depth in the object chain. Assigning `parent` sets `depth = parent.depth + 1`.

### parent

The parent `JSONobject` or `JSONarray`. Assigning a parent also copies `parent.defaultPropertyName` into the array.

```asp
arr.parent = jsonObject
Response.Write arr.depth
```

### defaultPropertyName

Inherited from the parent when `parent` is assigned. Defaults to `"data"` for a standalone array. It does not add a wrapper to the array's own JSON output; a serialized array is always a bare `[...]`.

```asp
arr.defaultPropertyName = "items"
```

## Methods

### Push

Appends a value to the array.

```asp
arr.Push value
```

**Parameters**

| Name    | Type      | Description                                                         |
| ------- | --------- | ------------------------------------------------------------------- |
| `value` | `Variant` | Value to append. Objects are stored by reference; scalars by value. |

The array grows automatically. Values can be scalars, `JSONobject`, `JSONarray`, or VBScript arrays.

```asp
Dim arr, nested
Set arr = New JSONarray

arr.Push "text"
arr.Push 42
arr.Push True
arr.Push Null

Set nested = New JSONobject
nested.Add "id", 1
arr.Push nested

arr.Write()
' ["text",42,true,null,{"id":1}]
```

### ItemAt

Returns the item at a zero-based index. `ItemAt` is the default function of the class, so `arr(index)` is equivalent.

```asp
value = arr.ItemAt(index)
value = arr(index)
```

**Parameters**

| Name    | Type      | Description            |
| ------- | --------- | ---------------------- |
| `index` | `Integer` | Zero-based item index. |

**Returns:** the item. Object items must be assigned with `Set`:

```asp
Dim child
Set child = arr(4)
child.Write()
```

**Errors:** if the array is empty or `index` is outside `0 .. length - 1`, error `9` (`JSON_ERROR_INDEX_OUT_OF_BOUNDS`) is raised.

### LoadRecordSet

Appends every remaining record of an ADO recordset to the array.

```asp
arr.LoadRecordSet rs
```

**Parameters**

| Name | Type              | Description     |
| ---- | ----------------- | --------------- |
| `rs` | `ADODB.Recordset` | Open recordset. |

**Remarks**

- Reads from the current position to EOF and leaves the recordset at EOF.
- Each record becomes a `JSONobject` whose properties match the field names.
- Existing items are kept; new records are appended.
- Call `rs.MoveFirst` if you need to read the recordset again afterwards.

```asp
Set arr = New JSONarray
arr.LoadRecordSet rs
arr.Write()
' [{"id":1,"name":"Ada"},{"id":2,"name":"Grace"}]
```

> To produce `{"data":[...]}` instead, use [`JSONobject.LoadRecordSet`](jsonobject.md#loadrecordset).

### Serialize

Returns the JSON representation of the array.

```asp
jsonString = arr.Serialize()
```

**Remarks**

- Output is compact; no indentation or extra whitespace is added.
- Temporarily sets `Response.LCID` to `1033` for number formatting, then restores it.
- If the array has a `JSONobject` parent, that object's serializer is reused.
- Otherwise a temporary `JSONobject` is created for the serialization call.

```asp
arr.Serialize()
' ["text",42,true,null]
```

### Write

Writes `Serialize()` directly to the response.

```asp
arr.Write()
' equivalent to: Response.Write arr.Serialize()
```

## Example

```asp
<%@ Language="VBScript" LCID="1033" %>
<% Option Explicit %>
<!--#include file="jsonObject.class.asp" -->
<%
Dim arr, row, i
Set arr = New JSONarray

For i = 1 To 3
    Set row = New JSONobject
    row.Add "id", i
    row.Add "name", "Item " & i
    arr.Push row
Next

Response.ContentType = "application/json"
arr.Write()
' [{"id":1,"name":"Item 1"},{"id":2,"name":"Item 2"},{"id":3,"name":"Item 3"}]
%>
```

## Related

- [JSONobject API](jsonobject.md)
- [Serialization reference](../reference/serialization.md)
- [Error reference](../reference/errors.md)
