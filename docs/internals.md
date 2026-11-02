# Internals

This page describes how the library is put together. It is aimed at maintainers and at users who need to reason about performance or edge cases.

## File layout

All classes live in [`jsonObject.class.asp`](../jsonObject.class.asp):

| Class        | Role                                                        |
| ------------ | ----------------------------------------------------------- |
| `JSONobject` | Parser, object builder, serializer, and recordset loader.   |
| `JSONarray`  | Dynamic array with `Push`, `ItemAt`, and recordset loading. |
| `JSONscript` | Raw-text wrapper that bypasses JSON quoting.                |
| `JSONpair`   | One name/value pair inside a `JSONobject`.                  |

The file also defines the module-level constants:

```vb
const JSON_ROOT_KEY = "[[JSONroot]]"
const JSON_DEFAULT_PROPERTY_NAME = "data"
const JSON_SPECIAL_VALUES_REGEX = "^(?:(?:t(?:r(?:ue?)?)?)|...)$"
const JSON_UNICODE_CHARS_REGEX = "\u(\d{4})"
```

## Parser

`JSONobject.Parse` is an **iterative state machine**, not a recursive descent parser. It walks the input one character at a time and switches between these modes:

| Mode         | Purpose                                                                                |
| ------------ | -------------------------------------------------------------------------------------- |
| `init`       | Starts a root, object, or array; creates the corresponding class and links its parent. |
| `openKey`    | Waits for the opening quote of an object key.                                          |
| `closeKey`   | Collects characters until an unescaped closing quote.                                  |
| `preValue`   | Waits for the `:` separator.                                                           |
| `openValue`  | Determines whether the value is a string, number, array, object, or keyword.           |
| `closeValue` | Collects the value and decodes string escapes.                                         |
| `addValue`   | Converts the collected text and stores it in the current object or array.              |
| `next`       | Handles `,`, `]`, and `}` and chooses the next state.                                  |

Two references, `currentObject` and `currentArray`, plus the `openObject` / `openArray` counters, track where values belong. Depth values on the objects are used to decide whether a value belongs to the current object or to an enclosing array.

### Root handling

The parser starts with `key = JSON_ROOT_KEY` and `root = me`.

- **Object root:** the current `JSONobject` is cleared and repopulated in place; `Parse` returns it.
- **Array root:** a `JSONarray` is created and returned. It is also added to `me` under the internal `JSON_ROOT_KEY` name. Serialization rewrites that sentinel key to `defaultPropertyName`, which is why `json.Write()` after parsing an array produces `{"data":[...]}`.

`getProperty` has a matching fallback: when a requested name equals `defaultPropertyName` and no real property has that name, it returns the sentinel pair. That is what makes `json("data")` return the root array.

### String and value handling

- String values are collected by slicing the input between quotes, then decoded with `Replace` calls for `\n`, `\r`, `\t`, `\b`, and `\\`, followed by a regular expression pass for `\uXXXX`.
- Property names are collected the same way structurally, but **no escape decoding is applied**.
- Numeric tokens are converted with `CDbl` (inside an en-US LCID block), so all numbers become `Double`.
- Keywords are converted through `Eval` after being restricted to short keyword-like tokens. The intended tokens are `true`, `false`, `null`, and `undefined`. `Eval` is not applied to arbitrary JSON text, but malformed input can still raise a native VBScript error here.

### LCID handling

Both parsing and serialization save `Response.LCID`, set it to `1033`, and restore it at the end of the method. This keeps number formatting independent of the page locale. The restore happens on normal completion; if a runtime error escapes in the middle of parsing or serialization, the response LCID may remain `1033` until you set it again.

## Object storage

A `JSONobject` stores an array of `JSONpair` references:

```vb
dim i_properties        ' JSONpair()
dim i_properties_count  ' number of pairs in use
dim i_properties_capacity
```

When adding a pair, the capacity grows as:

```vb
redim preserve i_properties(i_properties_capacity * 1.2 + 1)
```

`Remove` shifts subsequent pairs down and shrinks the buffer when the count drops below half the capacity.

### Lookup cost

`getProperty` performs a **linear scan** through the pairs to find a name. Both `Add` (duplicate check) and `Value` use it. Building a single object with *n* properties is therefore O(n²), and parsing a flat object with many keys has the same characteristic. This is usually fine for typical payloads but matters for very wide objects (thousands of keys).

`pairs` returns a trimmed copy of the internal array, so it is safe to iterate while inspecting the object.

## Array storage

`JSONarray` mirrors the object storage model:

```vb
dim i_items             ' Variant()
dim i_items_count
dim i_items_capacity
```

`Push` grows the buffer by roughly 20%:

```vb
redim preserve i_items(i_items_capacity * 1.2 + 1)
```

`ItemAt` is a direct index access, so indexed loops are fast. The `items` getter returns a trimmed copy of the buffer.

When `parent` is assigned to a `JSONarray`, the array copies the parent's `defaultPropertyName` and sets `depth = parent.depth + 1`.

## Serialization

Serialization is **recursive**:

- `JSONobject.SerializeObject` walks the `pairs` array and dispatches each value.
- `JSONarray.Serialize` delegates to `SerializeArray` on the parent object, or on a temporary `JSONobject` when there is no parent.
- `SerializeArray` handles both `JSONarray` and VBScript arrays; it detects the number of dimensions with `NumDimensions`.

The output is assembled by string concatenation and returned as one string. The entire document therefore exists in memory at least twice during `Serialize()` (the structure and the resulting string). For large documents, plan for memory proportional to the payload and adjust the ASP buffering limit and script timeout if needed.

Serialization is compact by design; there is no indentation option.

## ADO integration

`JSONobject.LoadRecordSet`:

1. Creates a new `JSONarray`.
2. For each record until EOF, creates a `JSONobject` and adds every field with `obj.Add field.name, field.value`.
3. Pushes the object into the array.
4. Calls `change defaultPropertyName, arr` to store the array, replacing any existing value under that name.

`JSONarray.LoadRecordSet` does the same per-record loop but pushes directly into the existing array.

`LoadFirstRecord` adds the current row's fields directly to the object.

`GetTypeName` wraps `TypeName` and supplies fallbacks for two database-specific variants that VBScript does not name correctly:

| VarType | Returned name |
| ------- | ------------- |
| `14`    | `Decimal`     |
| `16`    | `Integer`     |

`Add` also special-cases ADO `Field` objects, unwrapping `.Value`, and treats `IStringList` values as scalars.

## Error and cleanup behavior

- Custom errors use `Err.Raise` with codes `1`, `2`, `4`, `5`, and `9`. Error `3` is deprecated and unused.
- `NumDimensions` and `GetTypeName` use `On Error Resume Next` temporarily and restore error handling with `On Error GoTo 0`.
- `class_terminate` clears object references in the internal arrays to release COM references promptly. This is important for long-running pages or large object graphs.

## Known limitations

- **Not a strict validator.** Malformed input can raise native VBScript errors or produce partial/unexpected objects.
- **Object/array roots only.** Scalar JSON roots are not supported.
- **Property names are not escaped or unescaped.** Avoid quotes, backslashes, newlines, and control characters in keys.
- **Some escapes are not decoded.** `\"` remains escaped in parsed values; `\/` is not converted.
- **Two array dimensions maximum** for VBScript arrays.
- **Case-sensitive key lookup** under the default `Option Compare Binary`.
- **Duplicate keys raise error `2`** instead of taking the last value.
- **No pretty printing.** Output is always compact.
- **Non-standard values.** `Nothing` serializes as `undefined`; parsed `undefined` becomes `Empty` (and then `""`).
- **Dates have no timezone or milliseconds.**
- **In-memory only.** There is no streaming/SAX parser or writer.
- **Per-request use.** Instances hold parse state and parent references; do not share them across requests or store them in `Application` state.

## Tests

[`test.asp`](../test.asp) is the integration test page. It exercises:

- Parsing objects and arrays, including nested objects, empty containers, booleans, multiline strings, and Unicode escapes.
- `Add`, `Remove`, `Value`, and `Change`.
- `defaultPropertyName` changes.
- `JSONarray.Push`, iteration, and index access.
- `JSONobject.LoadRecordSet` and `JSONarray.LoadRecordSet` with an in-memory ADO recordset.
- `JSONscript` raw output.

Flags near the top of the page (`testLoad`, `testAdd`, `testRemove`, etc.) enable or disable sections. The page uses `Response.Buffer = True` and `Server.ScriptTimeout = 10`; the repository's [`web.config`](../web.config) raises the ASP buffering and timeout limits for larger tests.

## See also

- [Usage guide](guide.md)
- [Serialization reference](reference/serialization.md)
- [Error reference](reference/errors.md)
