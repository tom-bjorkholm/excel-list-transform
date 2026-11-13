# For future consideration

## Handling of empty parts and end separators when splitting

Rows (`s01_split_rows`) and columns (`s03_split_columns`) get different
parameters, because a row split splits at every separator while a column
split splits at a single separator (`maxsplit=1`). Dropping parts after a
column split does not give what users expect: `"a,b,"` split at the
rightmost `","` gives `"a,b"` and `""`, and split at the leftmost `","` it
gives `"a"` and `"b,"`. For columns, stripping separators at the ends of
the value before splitting is what makes sense.

### Common definitions

- A part is the text between two separators, or between a separator and
  the start or end of the value.
- A part is empty if it is `""` or consists only of whitespace.
- `in_excel_values_strip` is applied when reading the input, before any
  other processing. It applies only to Excel input, so with CSV input
  a value may still have whitespace at its start or end. Whitespace-only
  parts count as empty, so this does not affect which parts are empty.

### Row splitting: `empty_parts`

New mandatory key `empty_parts` in each rule in `s01_split_rows`.

```json
"empty_parts": "DROP_TRAILING"
```

Enum class `EmptyParts` with members:

- `KEEP_ALL`: keep all empty parts (the current behavior).
  `"a,,b,"` split at `","` gives `"a"`, `""`, `"b"` and `""`.
- `DROP_ALL`: drop all empty parts. `"a,,b,"` gives `"a"` and `"b"`.
- `DROP_TRAILING`: drop only the empty parts at the end.
  `"a,,b,"` gives `"a"`, `""` and `"b"`. `"a,b,,"` gives `"a"` and `"b"`.
  This is deliberately one-sided: it is common to write a separator after
  every element of a list (including the last one), but not before the
  first element.

If all parts are dropped (for example the value `","` or `""` with
`DROP_ALL`), the row is kept with the value `""`. (Today a value with one
or no parts returns the original row unchanged, so `","` would stay `","`.
This must be handled explicitly.)

### Column splitting: `end_separators`

New mandatory key `end_separators` in each rule in `s03_split_columns`.

```json
"end_separators": "STRIP_TRAILING"
```

Enum class `EndSeparators` with members:

- `KEEP`: keep the value as it is (the current behavior).
- `STRIP_LEADING`: remove separators and empty parts at the start of the
  value before splitting.
- `STRIP_TRAILING`: remove separators and empty parts at the end of the
  value before splitting.
- `STRIP_BOTH_ENDS`: do both.

All consecutive separators at the stripped end are removed, together with
any empty parts between them. Defined in terms of parts: split the value
at every separator, then `STRIP_TRAILING` cuts the value directly after
the last non-empty part, and `STRIP_LEADING` cuts it directly before the
first non-empty part. Examples with separator `","`:

- `"a,b,"` with `STRIP_TRAILING` becomes `"a,b"`.
- `"a,b,,"` and `"a,b, ,"` with `STRIP_TRAILING` become `"a,b"`.
- `",a,b"` with `STRIP_LEADING` becomes `"a,b"`.
- A value consisting only of separators (and whitespace) becomes `""`.

After stripping, the value is split according to `where` and stored
according to `store_single` / `right_name` as today.

### Naming rationale

- The key names what is handled (`empty_parts`, `end_separators`), the
  enum member names the action and its scope (`DROP_TRAILING`,
  `STRIP_TRAILING`).
- `LEADING` / `TRAILING` instead of `LEFT` / `RIGHT`, because `SplitWhere`
  already uses `LEFTMOST` / `RIGHTMOST` for another meaning.
- `KEEP` instead of `NONE`, to avoid confusion with JSON `null`.
- `string_to_enum_best_match` accepts unique prefixes, so for example
  `"STRIP_T"` works, while `"DROP"` or `"STRIP"` are rejected as ambiguous.

### Backward compatibility

Both keys are mandatory. For old configuration files, `ConfigReadOld`
(ROCF) adds the values that keep the old behavior:

- `('s01_split_rows', '[', 'empty_parts')`: `KEEP_ALL`
- `('s03_split_columns', '[', 'end_separators')`: `KEEP`

Both enums need entries in `parse_converters()`, validation in
`_split_rows_step()` / `_split_columns_step()`, descriptions in
`edit_descriptions.py`, and updates in `generate_cfg.py` /
`generate_txt.py`.
