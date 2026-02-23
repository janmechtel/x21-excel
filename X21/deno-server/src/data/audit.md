**COMPREHENSIVE EXCEL Sheet AUDIT PROMPT**

Conduct a thorough audit of this sheet and identify all errors,
categorized by severity level.

**AUDIT SCOPE:**

1. **Formula Integrity**
   - Check for broken formulas, #REF!, #VALUE!, #DIV/0!, #N/A, #NUM! errors
   - Identify circular references (both intentional and unintentional)
   - Detect hardcoded values within formulas where references should be used
   - Flag inconsistent formulas across rows/columns where they should be uniform
   - Verify array formulas are working correctly

2. **Calculation Logic**
   - Validate mathematical accuracy of key calculations
   - Check aggregation formulas (SUM, SUBTOTAL) for missing or extra cells
   - Verify percentage calculations and rate applications
   - Confirm period-over-period calculations (growth rates, changes)
   - Check that precedents and dependents flow logically

3. **Structural Integrity**
   - Identify merged cells that may disrupt formulas
   - Flag hidden rows/columns that contain critical formulas or assumptions
   - Check for inconsistent formatting that may hide errors
   - Verify named ranges are correctly defined and used
   - Check for broken links to external workbooks

4. **Financial/Business Logic**
   - Verify accounting principles (debits/credits, balance sheet balancing)
   - Check cash flow calculations and reconciliations
   - Validate depreciation and amortization schedules
   - Confirm tax calculations and carryforwards
   - Verify debt schedules (principal, interest, balance calculations)
   - Check working capital calculations

5. **Data Consistency**
   - Compare linked cells for consistency across sheets
   - Verify input assumptions are used consistently throughout
   - Check date ranges and period consistency
   - Flag mixed currencies or unit inconsistencies
   - Identify duplicate data or redundant calculations

6. **Formatting & Presentation**
   - Flag unlabeled sections or unclear headers
   - Identify inputs that aren't clearly distinguished from calculations
   - Check for inconsistent number formatting
   - Verify units are clearly labeled (000s, millions, percentages)

**SEVERITY CLASSIFICATION:**

**CRITICAL ERRORS** (Stop work - immediate fix required):

- Formulas producing error values in key outputs
- Broken links affecting calculations
- Balance sheet doesn't balance
- Cash flow calculations fundamentally broken
- Circular references causing incorrect results
- Hardcoded values overriding key formulas

**SEVERE ERRORS** (High priority - affects decision-making):

- Material calculation errors in financial statements
- Inconsistent formulas in summation rows
- Missing or incorrect growth rate applications
- Tax calculation errors
- Debt schedule miscalculations
- Wrong cells referenced in key metrics

**MODERATE ERRORS** (Should fix - may affect accuracy):

- Minor formula inconsistencies
- Formatting issues that obscure data
- Inefficient formula construction
- Non-critical naming inconsistencies
- Minor linking issues

**LOW PRIORITY** (Cosmetic/efficiency):

- Formatting inconsistencies
- Optimization opportunities
- Documentation gaps

**OUTPUT FORMAT:**

For each error identified, provide:

1. **Issue**: Clear description of what's wrong
2. **Location**: Sheet name and cell reference, for multiple ranges repeat the sheet name for clarity so 'Sheet1!A1:A10',  'Sheet1 C1:C10'
3. **Severity**: Critical/Severe/Moderate - skip Low Priority unless requested
4. **Impact**: What this affects in the model
5. **Current Formula/Value**: What's currently there
6. **Recommendation**: Specific fix needed

Prioritize the audit by focusing first on output sheets and working backwards to
inputs, and prioritize items that flow into key decision metrics (IRR, NPV,
multiples, returns, valuations).

Begin with a summary of findings showing count of errors by severity, then
provide detailed findings.

Don't use the formatting tool for this unless the user is explicitly asking for auditing around formatting issues.
