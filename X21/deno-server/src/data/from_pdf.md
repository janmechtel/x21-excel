# Recreate These Attachments in Excel

## Extraction Guidelines

- When extracting/replicating a PDF default to creating a new sheet,
  if filling a template then fill it in the current sheet, if in doubt
  ask the user where to insert the data, default to new sheet
- Only do number formatting, do NOT DO formatting of headers,
  background etc. unless specifically requested.
- **CRITICAL: EXACT Number Format Preservation** - When recreating data from
  attachments, you MUST preserve the EXACT number format from the source
  document. This is a MANDATORY requirement that must be followed precisely:

  **Before writing any numbers, carefully observe the source document:**
  1. **Negative numbers**: Look at how negatives are displayed in the PDF/image:
     - If source shows (123) or (1,234.56), use parentheses format: `#,##0;(#,##0)`
     - If source shows -123 or -1,234.56, use minus format: `#,##0;-#,##0`
     - DO NOT assume - always check the source first
  2. **Currency symbols**: Observe the exact currency symbol and placement:
     - If source shows $1,234, use `$#,##0` (dollar before)
     - If source shows 1,234€ or 1,234 €, use `#,##0"€"` or `#,##0 "€"` (euro after)
     - If source shows £1,234, use `£#,##0` (pound before)
     - If source shows 1,234 USD, use `#,##0" USD"` (text after)
     - Match the EXACT symbol, placement, and spacing from the source
  3. **Decimal places**: Match the exact number of decimal places shown:
     - If source shows 1,234.56, use 2 decimals
     - If source shows 1,234, use 0 decimals
     - If source shows 1,234.5, use 1 decimal
  4. **Thousands separators**: Match the separator style:
     - If source shows 1,234 (comma), use comma separator
     - If source shows 1.234 (period), use period separator
     - If source shows 1 234 (space), use space separator
  5. **Percentage formatting**: Match exactly:
     - If source shows 12.5%, use `0.0%`
     - If source shows 12.50%, use `0.00%`
     - If source shows 12%, use `0%`
  6. **Combined formats**: For currency with negatives, combine correctly:
     - Source: $(123) → Format: `$#,##0;($#,##0)`
     - Source: -$123 → Format: `$#,##0;-$#,##0`
     - Source: (€1,234) → Format: `#,##0"€";(#,##0"€")`

  **Validation step**: After writing, verify that the displayed format in Excel
  matches the source document exactly. If it doesn't match, correct it immediately.
- stay very close to the original layout
- Write in bulk: use write_values_batch for all writes. One operation
  per contiguous block; multiple operations for disjoint ranges. Emit
  exactly ONE write_values_batch call per response; avoid row-by-row
  tool calls.
- use formulas where possible to do sums/subtotals instead of
  hardcoded values
- Prefer SUM if the range of the cells is adjacent so =SUM(A1:A2)
  instead of =A1+A2, But use + if the cells are further apart like
  =A1+A10. Avoid subtotals unless specifically requested
- When asked to extract certain information like the name from a CV
  or the starting date for a contract, create a separate sheet called
  "extracted_data" and put the requested information in clearly
  labeled cells + record their source (eg "from page 2, line 5")
- Clarify with the user if there is no image/pdf attachment or no
  data visible in th additional user prompt below

## Delta Check Columns

- create Delta check columns on all the rows with sums put the values
  from the attachment file into separate columns to the far right
- to see if there is any delta compared to the calculated subtotals
- So if the calculated value is in A13 =Sum(A2:A12) then for example
  put the delta in D13 =15-A13 where 15 is the value from the
  attachmentfile
- Put the checksum columns on the same sheet right to the main area
- Repeat the delta columns in the same order with the same labels as
  the original columns
- have checksums in all cells where there original values have been
  replace by calculationsi(in cells without calculations, the
  checksums are meaningless). So the checksum cells should be on the
  same as the rows with sums/subtotals in the original file. That
  Makes it easier to review.
- highlight any deviations (display empty cells for 0) but use number
  formatting to show non-zero values in red

## Verification (MANDATORY)

**After writing data, you MUST:**

1. **Read back the written data** using read_values to verify what
   was actually written
2. **Check all formulas for errors** - scan for #REF!, #VALUE!,
   #DIV/0!, #N/A, #NUM!, #NAME?, #NULL! errors
3. **Fix any formula errors immediately** before proceeding
4. **Validate number formats match source** (CRITICAL):
   - Compare the displayed format in Excel with the source document
   - Verify negative numbers match (parentheses vs minus)
   - Verify currency symbols and placement match exactly
   - Verify decimal places match
   - Verify thousands separators match
   - If formats don't match, correct the numberFormat immediately
5. **Validate checksums and totals**:
   - Verify all sum formulas match expected totals from the source
     document
   - Check that subtotals add up correctly to grand totals
   - For income statement type files, ensure all subtotals and totals
     match perfectly (e.g., net result should match exactly)
6. **Review delta check columns**:
   - Check if very small deviations are explained via invisible
     rounding of sums and give hints if that could be the case
   - Be extra careful with the positive / negative sign for the delta
     check as it might indicate a formula mistake with regards to the
     plus/minus sign especially if the delta equals 2x the amount
   - Fix any significant discrepancies automatically
7. **Report findings**:
   - Inform the user of any errors found and fixed
   - Ask the user if you should take another pass trying to fix
     remaining deviations
   - Highlight any issues that couldn't be automatically resolved

**This validation step is MANDATORY and must happen after every data
extraction. Do not skip validation.**
