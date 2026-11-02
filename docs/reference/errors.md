# Error reference

The library raises errors with `Err.Raise`, using the constants defined at the top of `jsonObject.class.asp`.

## Error constants

| Constant | Value | Raised by | Meaning |
| --- | --- | --- | --- |
| `JSON_ERROR_PARSE` | `1` | `JSONobject.Parse` | The input is too short to be valid JSON (fewer than 2 characters after trimming). |
| `JSON_ERROR_PROPERTY_ALREADY_EXISTS` | `2` | `JSONobject.Add`, `JSONobject.Parse`, `JSONobject.LoadFirstRecord` | A property with the same name already exists, a JSON object contains duplicate keys, or a record field conflicts with an existing property. |
| `JSON_ERROR_PROPERTY_DOES_NOT_EXISTS` | `3` | — | **Deprecated.** Reserved for compatibility; not raised by the current code. |
| `JSON_ERROR_NOT_AN_ARRAY` | `4` | `JSONarray.items` (setter) | A non-array value was assigned to `JSONarray.items`. |
| `JSON_ERROR_NOT_A_STRING` | `5` | `JSONscript.value` (setter) | A non-string value was assigned to `JSONscript.value`. |
| `JSON_ERROR_INDEX_OUT_OF_BOUNDS` | `9` | `JSONarray.ItemAt` | The index is negative, greater than or equal to `length`, or the array is empty. The number matches VBScript's "Subscript out of range" error. |

`Err.Source` is set to the class name (`JSONobject`, `JSONarray`, or `JSONscript`).

## Handling errors

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

### Parse errors

```asp
On Error Resume Next
json.Parse "{"
If Err.Number = JSON_ERROR_PARSE Then
    Response.Write "The input is too short."
End If
Err.Clear
On Error GoTo 0
```

Only a short input raises the custom parse error. Malformed JSON may instead raise a **native VBScript error** (for example a type mismatch or subscript error) because the parser is not a strict validator. Always treat parsing untrusted input as a potentially failing operation.

### Duplicate properties

```asp
On Error Resume Next
json.Add "name", "Ada"
json.Add "name", "Grace"        ' Err.Number = 2
If Err.Number = JSON_ERROR_PROPERTY_ALREADY_EXISTS Then
    json.Change "name", "Grace" ' update instead
    Err.Clear
End If
On Error GoTo 0
```

### Array index errors

```asp
On Error Resume Next
Dim value
value = arr(10)
If Err.Number = JSON_ERROR_INDEX_OUT_OF_BOUNDS Then
    Response.Write "No such item."
    Err.Clear
End If
On Error GoTo 0
```

Check `length` first when an out-of-range index is expected:

```asp
If index >= 0 And index < arr.length Then
    value = arr(index)
Else
    value = Null
End If
```

## Common non-custom errors

| Error | Typical cause |
| --- | --- |
| `Object required` / `Object doesn't support this property or method` | The include file was not loaded, an object variable was used without `Set`, or a non-object was treated as one. |
| `Subscript out of range` (`9`) | An out-of-range array access, including `JSONarray.ItemAt`. |
| `Type mismatch` | An unexpected value type was serialized or assigned, or a malformed JSON string confused the parser. |
| Custom error `4` | Assigning a non-array to `JSONarray.items`. |
| Custom error `5` | Assigning a non-string to `JSONscript.value`. |

## See also

- [JSONobject API](../api/jsonobject.md)
- [JSONarray API](../api/jsonarray.md)
- [Internals](../internals.md)
