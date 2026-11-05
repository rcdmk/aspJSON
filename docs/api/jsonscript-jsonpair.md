# JSONscript and JSONpair API

These supporting classes are defined in the same file as `JSONobject` and `JSONarray`.

## JSONscript

Wraps a raw string that should be written to JSON output without quoting or escaping. The typical use is embedding a JavaScript expression or callback:

```asp
Dim script, json
Set script = New JSONscript
script.value = "function () { alert('test'); }"

Set json = New JSONobject
json.Add "callback", script

json.Write()
' {"callback":function () { alert('test'); }}
```

The class version in the source is `1.0.0` (there is no public `version` property).

### value

```asp
script.value = "raw text"
```

- Get/set property.
- Must be a `String`; any other type raises `JSON_ERROR_NOT_A_STRING` (`5`).
- An empty string is replaced with the literal text `null`.

```asp
script.value = ""                   ' serializes as: null
```

### Serialize

```asp
text = script.Serialize()
```

Returns the stored value exactly as-is. No validation or escaping is performed.

### Write

```asp
script.Write()
' equivalent to: Response.Write script.Serialize()
```

### Important notes

- `JSONscript` performs **no validation**. Only use it with trusted, well-formed content; otherwise it can produce invalid JSON or inject arbitrary output.
- It is handled specially only when it is a **property value of a `JSONobject`**. If a `JSONscript` is placed inside a `JSONarray`, the serializer treats it as an ordinary value and quotes it as a string.
- An empty `JSONscript` produces the bare keyword `null`, not a quoted string.

## JSONpair

Represents one name/value pair inside a `JSONobject`. `JSONobject.pairs` exposes a snapshot array of these objects.

| Member   | Access  | Type         | Description                                                          |
| -------- | ------- | ------------ | -------------------------------------------------------------------- |
| `name`   | get/set | `String`     | The pair's key.                                                      |
| `value`  | get/set | `Variant`    | The pair's value. Use `Set` when reading or writing an object value. |
| `parent` | get/set | `JSONobject` | The object that owns the pair.                                       |

```asp
Dim pair, value
For Each pair In json.pairs
    If IsObject(pair.value) Then
        Set value = pair.value
        Response.Write pair.name & ": " & value.Serialize()
    Else
        Response.Write pair.name & ": " & pair.value
    End If
Next
```

`JSONpair` is an implementation detail. Prefer `Add`, `Change`, `Remove`, and `Value` on `JSONobject` instead of manipulating pairs directly.

## Related

- [JSONobject API](jsonobject.md)
- [JSONarray API](jsonarray.md)
