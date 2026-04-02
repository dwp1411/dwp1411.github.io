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

## Selective Column Filtering

### Problem
You want to pull data from a source sheet into columns A, B, and E of a destination sheet, but keep columns C and D empty for manual entry. A single array formula in A2 (like `={...}`) would overwrite C and D or cause a `#REF!` error.

### Solution
Use separate `FILTER` formulas in the top cell of each column you want to populate (e.g., A2, B2, E2). This leaves the intermediate columns (C, D) free for manual data.

**Example Setup:**
*   Source Sheet: `'Overall Staff'`
*   Filter Criteria: Rows where Column B contains "Abel Montoya"

**Formulas:**

**Cell A2 (Name):**
```excel
=IFNA(FILTER('Overall Staff'!A:A, ISNUMBER(SEARCH("Abel Montoya", 'Overall Staff'!B:B))), "No Records Found")
```

**Cell B2 (Supervisor):**
```excel
=IFNA(FILTER('Overall Staff'!B:B, ISNUMBER(SEARCH("Abel Montoya", 'Overall Staff'!B:B))), "")
```

**Cell E2 (Temp Y/N):**
```excel
=IFNA(FILTER('Overall Staff'!E:E, ISNUMBER(SEARCH("Abel Montoya", 'Overall Staff'!B:B))), "")
```
