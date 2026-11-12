# For future consideration

## Optional `empty_parts` parameter for splitting rows and columns

Add an optional parameter `empty_parts` to the rules in `s01_split_rows`
(splitting rows) and `s03_split_columns` (splitting columns). It controls
how empty parts are handled. Empty parts come from a separator at the start
or end of the value, or from two adjacent separators.

Possible values:

- `keep_all`: keep all empty parts. This is the current behavior for
  splitting rows. For example `"a,,b,"` split at `","` gives
  `"a"`, `""`, `"b"` and `""`.
- `drop_all`: drop all empty parts. `"a,,b,"` gives `"a"` and `"b"`.
- `drop_trailing`: drop only trailing empty parts. `"a,,b,"` gives
  `"a"`, `""` and `"b"`.

Open questions:

- Which value should be the default? `keep_all` keeps the current
  behavior and backward compatibility.
- What should happen when a value consists only of separators and all
  parts are dropped: keep the row with an empty value, or drop the row?
- How does `empty_parts` combine with `where` and `store_single` /
  `right_name` when splitting columns? A column split gives at most two
  parts.
