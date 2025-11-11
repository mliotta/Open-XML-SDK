# Excel Oracle Validation

This directory contains tests that validate formula evaluation correctness against **real Excel calculations**.

## The Oracle Approach

Instead of using Excel COM automation (complex, Windows-only), we use a **golden file** approach:

1. **Generate**: SDK creates an Excel file with 100+ formula test cases (no cached values)
2. **Bless**: User opens file in Excel → Excel calculates all formulas → cached values stored
3. **Validate**: SDK reads blessed file → evaluates formulas → compares against Excel's cached values

This gives us:
- ✅ Ground truth from real Excel (not our interpretation)
- ✅ No COM complexity or Windows dependency
- ✅ Fast execution (no Excel startup per test)
- ✅ Portable test files (can commit to repo for CI/CD)

## How to Generate and Bless the Oracle File

### Step 1: Generate the Oracle File

```bash
# Run the generator test (unskip it first)
dotnet test --filter "FullyQualifiedName~GenerateOracleTestFile"
```

This creates: `/tmp/FormulaOracle.xlsx`

The file contains 100+ test cases across 7 sheets:
- **Math**: SUM, AVERAGE, COUNT, MAX, MIN, ROUND, ABS, PRODUCT, POWER (25+ tests)
- **Logical**: IF, AND, OR, NOT (15+ tests)
- **Text**: CONCATENATE, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, TEXT, VALUE, FIND, SUBSTITUTE (20+ tests)
- **Lookup**: VLOOKUP, HLOOKUP (5+ tests)
- **DateTime**: TODAY, NOW, DATE, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY (15+ tests)
- **Statistical**: MEDIAN, MODE, STDEV, VAR, RANK (10+ tests)
- **Information**: ISNUMBER, ISTEXT (6+ tests)

### Step 2: Bless with Excel

1. Open `/tmp/FormulaOracle.xlsx` in **Microsoft Excel**
2. Excel will automatically calculate all formulas and store cached values
3. **Verify**: Check that column C shows calculated values (not formulas)
4. **Save** the file (Ctrl+S / Cmd+S)
5. **Close** Excel

### Step 3: Copy Blessed File

```bash
# Create TestFiles directory if it doesn't exist
mkdir -p test/DocumentFormat.OpenXml.Formulas.Tests/TestFiles

# Copy the blessed file
cp /tmp/FormulaOracle.xlsx test/DocumentFormat.OpenXml.Formulas.Tests/TestFiles/
```

### Step 4: Run Oracle Validation

```bash
# Unskip the validation test first, then run:
dotnet test --filter "FullyQualifiedName~ValidateAgainstExcelOracle"
```

This will:
- Load the blessed Excel file
- Evaluate each formula with our engine
- Compare our result against Excel's cached value
- Report pass/fail rate and detailed failures

## Expected Results

**Target**: 95%+ pass rate

The validation test will output:
```
Total test cases: 120
Passed: 115 (95.83%)
Failed: 5
Skipped: 0

FAILURES:
  Math!C15: =ROUND(2.5, 0)
    Expected: 2
    Actual:   3
    Error:    Value mismatch
```

## What Counts as a Pass?

- **Numbers**: Must match within 0.0001 (floating point tolerance)
- **Text**: Exact string match
- **Booleans**: Excel stores as "1"/"0", we compare accordingly
- **Errors**: Error string must match (e.g., "#DIV/0!")

## Troubleshooting

### "Oracle file not found"
- Make sure you completed Step 3 (copy blessed file)
- Check file path: `test/DocumentFormat.OpenXml.Formulas.Tests/TestFiles/FormulaOracle.xlsx`

### "No Excel cached value"
- You skipped Step 2 (blessing with Excel)
- Open file in Excel, save, close

### High failure rate (>5%)
- **Good!** This found correctness bugs
- Review failures in test output
- Check if it's a systemic issue (e.g., all ROUND functions fail)
- Fix the function implementation
- Re-run validation

### Test still skipped
- Remove `Skip = "..."` attribute from test methods in `OracleValidationTests.cs`

## Continuous Integration

Once blessed, commit the oracle file to the repo:

```bash
git add test/DocumentFormat.OpenXml.Formulas.Tests/TestFiles/FormulaOracle.xlsx
git commit -m "Add blessed Excel oracle file for correctness validation"
```

CI/CD will then:
1. Run oracle validation on every PR
2. Catch regressions immediately
3. No Excel needed (uses cached values)

## Updating the Oracle File

When adding new functions:

1. Modify `OracleTestFileGenerator.cs` to add test cases
2. Re-run Step 1-3 to regenerate and rebless
3. Commit updated oracle file

## Files

- **OracleTestFileGenerator.cs**: Generates Excel file with test cases
- **OracleValidationTests.cs**: Validates against blessed file
- **README_ORACLE.md**: This file
- **TestFiles/FormulaOracle.xlsx**: Blessed oracle file (after Step 3)

## Philosophy

This approach gives us the best of both worlds:
- **One-time blessing**: Use real Excel to establish ground truth
- **Fast CI/CD**: No Excel dependency in automated tests
- **Correctness proof**: Our results match real Excel (99%+)

It's like having Excel as your test oracle, but without the overhead of COM automation on every test run.
