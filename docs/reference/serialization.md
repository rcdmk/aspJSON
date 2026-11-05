# Serialization reference

This page describes exactly how `aspJSON` converts between JSON text and VBScript values.

## Serialization

`JSONobject.Serialize()`, `JSONobject.Write()`, `JSONarray.Serialize()`, and `JSONarray.Write()` produce **compact** JSON: no indentation, newlines, or spaces between tokens.

While serializing, the library temporarily sets `Response.LCID = 1033` (en-US) so numbers and dates use invariant formatting, then restores the previous LCID.

### Value mapping

| VBScript value                                                       | JSON output                                                       | Notes                                                                                                       |
| -------------------------------------------------------------------- | ----------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| `Null`                                                               | `null`                                                            | Standard JSON null.                                                                                         |
| `Nothing`                                                            | `undefined` only via `SerializeValue`; `Serialize()` raises `424` | `IsObject(Nothing)` is true, so normal object/array serialization fails with `Object required`. Use `Null`. |
| `Empty`                                                              | `""`                                                              | Serialized as an empty quoted string.                                                                       |
| `True` / `False`                                                     | `true` / `false`                                                  |                                                                                                             |
| `Byte`, `Integer`, `Long`, `Single`, `Double`, `Currency`, `Decimal` | number                                                            | Unquoted. Decimal separator is `.`.                                                                         |
| `Date`                                                               | `"YYYY-MM-DDTHH:MM:SS"`                                           | Server-local time; no milliseconds or timezone offset.                                                      |
| `String`, `Char`, `Empty`                                            | `"..."`                                                           | Escaped; see below.                                                                                         |
| `JSONobject`                                                         | object                                                            | Recursively serialized.                                                                                     |
| `JSONarray`                                                          | array                                                             | Recursively serialized.                                                                                     |
| `JSONscript`                                                         | raw text                                                          | Only when used as a `JSONobject` property value. Not quoted or escaped.                                     |
| One-dimensional VBScript array                                       | array                                                             | Elements serialized recursively.                                                                            |
| Two-dimensional VBScript array                                       | array of arrays                                                   | Outer array by first dimension, inner arrays by second dimension.                                           |
| Other non-object types                                               | `"TypeName"`                                                      | Fallback: the VBScript type name as a quoted string.                                                        |

### Numbers

- Numeric values are written using `1033` locale formatting, so the decimal separator is always `.`.
- Values are stored as VBScript `Double` when parsed. Extremely large integers or values needing more than about 15–17 significant digits can lose precision.
- There is no special handling for `NaN` or infinities; avoid those values.

### Dates

`Date` values become ISO 8601-like strings using the server's local date and time:

```asp
json.Add "created", Now
' "created":"2026-11-02T14:05:09"
```

There is no timezone designator, no milliseconds, and no `Z` suffix. If consumers need UTC, convert the value to UTC before adding it, or store a preformatted string.

Parsed date-looking strings stay strings; `Parse` does not convert them to `Date` values.

### String escaping

`EscapeCharacters` performs these replacements, in order:

| Character                    | Output |
| ---------------------------- | ------ |
| `\` (backslash)              | `\\`   |
| `"` (double quote)           | `\"`   |
| `CR` (`vbCr`, U+000D)        | `\r`   |
| `LF` (`vbLf`, U+000A)        | `\n`   |
| `TAB` (`vbTab`, U+0009)      | `\t`   |
| backspace (`vbBack`, U+0008) | `\b`   |

Not escaped:

- Forward slash `/`
- Form feed (`vbFormFeed`) and other control characters
- Non-ASCII characters (they are emitted as-is)

If your data can contain form feeds or arbitrary control characters, clean or pre-escape it yourself to guarantee strictly valid JSON. Non-ASCII output relies on the page's code page / response charset being set correctly (for example UTF-8).

> **Property names are not escaped.** `Serialize` writes keys as-is between double quotes. Avoid quotes, backslashes, newlines, and control characters in property names, or the result may not be valid JSON.

### Arrays

When a VBScript array is passed to `Add` or `Change`, it is converted to a `JSONarray`:

```asp
json.Add "values", Array(1, "two", True, Null)
' "values":[1,"two",true,null]
```

- One- and two-dimensional arrays are supported.
- A two-dimensional array is serialized as an array of arrays, looping first over dimension 1 and then dimension 2.
- Three or more dimensions are not supported. `Add` accepts them, but serialization raises error `9` (`Subscript out of range`), and object cleanup may raise the same error.
- Jagged/nested arrays work when built with nested `JSONarray` objects or nested VBScript arrays.

### Raw output

`JSONscript` bypasses escaping entirely:

```asp
Set script = New JSONscript
script.value = "function () { return 1; }"
json.Add "callback", script
' "callback":function () { return 1; }
```

Only use it with trusted content. See [JSONscript](../api/jsonscript-jsonpair.md#jsonscript).

## Parsing

`JSONobject.Parse()` walks the input with an iterative state machine and builds `JSONobject` / `JSONarray` instances.

### Parsing type mapping

| JSON input           | VBScript value       |
| -------------------- | -------------------- |
| `"text"`             | `String`             |
| `123`, `-1.5`, `1e3` | `Double`             |
| `true` / `false`     | `Boolean`            |
| `null`               | `Null`               |
| `undefined`          | `Empty` (via `Eval`) |
| `{ ... }`            | `JSONobject`         |
| `[ ... ]`            | `JSONarray`          |

Because `undefined` maps to `Empty`, and `Empty` serializes as `""`, an `undefined` token does not round-trip back to `undefined`.

### String value escapes

The parser decodes these escapes in **string values**:

| Input    | Decoded to                   |
| -------- | ---------------------------- |
| `\n`     | line feed                    |
| `\r`     | carriage return              |
| `\t`     | tab                          |
| `\b`     | backspace                    |
| `\\`     | backslash                    |
| `\uXXXX` | Unicode character via `ChrW` |

For characters outside the Basic Multilingual Plane, provide a surrogate pair (`\uD83D\uDE00` for 😀).

Other JSON escapes are not decoded. For example, `\"` is used to find the end of the string but the backslash remains in the parsed value, and `\/` remains as-is.

> **Property names are not unescaped.** Escape sequences are decoded only in string values. A key such as `"line\nbreak"` keeps the literal characters `\` and `n`.

### Other parsing behavior

- Surrounding whitespace is trimmed.
- The JSON root must be an object or an array. A bare scalar root is not supported.
- Empty objects (`{}`) and empty arrays (`[]`) are supported.
- Duplicate object keys raise `JSON_ERROR_PROPERTY_ALREADY_EXISTS` (`2`).
- Object keys are matched using VBScript string comparison (binary/case-sensitive by default).
- Numbers are parsed with `CDbl`; they become `Double`.
- Special keywords (`true`, `false`, `null`, `undefined`) must begin with a lowercase letter. A keyword starting with an uppercase letter, such as `TRUE`, is not recognized and can be silently skipped.
- The parser is not a strict RFC 8259 validator. Malformed input may raise a VBScript runtime error or produce an unexpected result. Validate untrusted input at the boundary if strict rejection is required.
- The parser temporarily sets `Response.LCID = 1033` and restores the previous LCID when it finishes.
- Input shorter than two characters raises `JSON_ERROR_PARSE` (`1`).

## See also

- [JSONobject API](../api/jsonobject.md)
- [JSONarray API](../api/jsonarray.md)
- [Error reference](errors.md)
