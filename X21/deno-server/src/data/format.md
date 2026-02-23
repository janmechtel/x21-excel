Format everything and make it look nice an consistent.

If the user only selected a single cell or small range, apply formatting to the entire table or data range or sheet. When in doubt, ask the user for clarification.

NEVER Justify / middle align text in cells, avoid centered labels undless explicitly requested.

# Color Coding Standards

Unless otherwise stated by the user or existing template

## Industry-Standard Color Conventions

- Blue text (RGB: 0,0,255): Hardcoded inputs, and numbers users will change for scenarios
- Black text (RGB: 0,0,0): ALL formulas and calculations
- Green text (RGB: 0,128,0): Links pulling from other worksheets within same workbook
- Red text (RGB: 255,0,0): External links to other files
- Yellow background (RGB: 255,255,0): Key assumptions needing attention or cells that need to be updated

# Number Formatting Standards

## Required Format Rules

- Years: Format as text strings (e.g., "2024" not "2,024")
- Currency: Use $#,##0 format; ALWAYS specify units in headers ("Revenue ($mm)")
- Zeros: Use number formatting to make all zeros "-", including percentages (e.g., "$#,##0;($#,##0);-")
- Percentages: Default to 0.0% format (one decimal)
- Multiples: Format as 0.0x for valuation multiples (EV/EBITDA, P/E)
- Negative numbers: Use parentheses (123) not minus -123
