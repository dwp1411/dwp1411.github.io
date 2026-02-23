# Google Sheets Tips

## FILTER function error: "mismatched range sizes"

### Problem
When using the `FILTER` function in Google Sheets, you might encounter an error like:
`FILTER has mismatched range sizes. Expected row count: 1001. column count: 1. Actual row count: 1, column count: 1.`

This happens if you provide a third argument to `FILTER` intended as a "no records found" message (similar to Excel's `FILTER` function).
Example causing error:
```excel
=FILTER('Overall Staff'!A:C, ISNUMBER(SEARCH("Abel Montoya", 'Overall Staff'!B:B)), "No Records Found")
```

### Explanation
Google Sheets' `FILTER` function syntax is `FILTER(range, condition1, [condition2, ...])`.
It interprets the third argument `"No Records Found"` as a second condition. Since this string is a single value (1 row) and your data range has many rows (e.g., 1001), the sizes do not match.

### Solution
Wrap the `FILTER` function in `IFNA` (or `IFERROR`) to handle the case where no data matches the condition.

Corrected formula:
```excel
=IFNA(FILTER('Overall Staff'!A:C, ISNUMBER(SEARCH("Abel Montoya", 'Overall Staff'!B:B))), "No Records Found")
```
