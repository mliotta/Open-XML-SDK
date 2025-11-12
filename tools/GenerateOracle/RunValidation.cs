using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Features.FormulaEvaluation;
using EvalCellValue = DocumentFormat.OpenXml.Features.FormulaEvaluation.CellValue;

public class RunValidation
{
    public static void Main()
    {
        var oracleFilePath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "../../test/DocumentFormat.OpenXml.Formulas.Tests/TestFiles/FormulaOracle.xlsx");

        if (!File.Exists(oracleFilePath))
        {
            Console.WriteLine($"ERROR: Oracle file not found at: {oracleFilePath}");
            return;
        }

        Console.WriteLine($"Validating against oracle file: {oracleFilePath}");
        Console.WriteLine();

        var results = new ValidationResults();

        using (var doc = SpreadsheetDocument.Open(oracleFilePath, false))
        {
            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.Get<IFormulaEvaluator>();
            if (evaluator == null)
            {
                Console.WriteLine("ERROR: Failed to get formula evaluator feature");
                return;
            }

            foreach (var worksheetPart in doc.WorkbookPart!.WorksheetParts)
            {
                var worksheet = worksheetPart.Worksheet;
                var sheetName = GetSheetName(doc.WorkbookPart, worksheetPart);

                Console.WriteLine($"Validating sheet: {sheetName}");

                // Column C contains formulas with Excel-calculated cached values
                var formulaCells = worksheet.Descendants<Cell>()
                    .Where(c => c.CellReference != null &&
                               c.CellReference.Value!.StartsWith("C") &&
                               c.CellFormula != null)
                    .OrderBy(c => GetRowNumber(c.CellReference!.Value!))
                    .ToList();

                foreach (var cell in formulaCells)
                {
                    results.Total++;

                    var cellRef = cell.CellReference!.Value!;
                    var rowNum = GetRowNumber(cellRef);

                    // Skip header row
                    if (rowNum == 1)
                    {
                        results.Skipped++;
                        continue;
                    }

                    var formula = cell.CellFormula!.Text;
                    var excelValue = cell.CellValue?.Text;

                    if (string.IsNullOrEmpty(excelValue))
                    {
                        results.Skipped++;
                        continue;
                    }

                    // Evaluate with our engine
                    var evalResult = evaluator.TryEvaluate(worksheet, cell);

                    if (!evalResult.IsSuccess)
                    {
                        results.Failed++;
                        Console.WriteLine($"  FAIL {cellRef}: {formula}");
                        Console.WriteLine($"       Error: {evalResult.Error?.Message}");
                        continue;
                    }

                    // Compare values
                    var ourValue = evalResult.Value;
                    bool match = CompareValues(ourValue, excelValue, cell.DataType?.Value);

                    if (match)
                    {
                        results.Passed++;
                    }
                    else
                    {
                        results.Failed++;
                        Console.WriteLine($"  FAIL {cellRef}: {formula}");
                        Console.WriteLine($"       Expected: {excelValue}");
                        Console.WriteLine($"       Got: {FormatValue(ourValue)}");
                    }
                }

                Console.WriteLine();
            }
        }

        // Report results
        Console.WriteLine("================================================================================");
        Console.WriteLine($"Total test cases: {results.Total}");
        Console.WriteLine($"Passed: {results.Passed} ({results.PassRate:P2})");
        Console.WriteLine($"Failed: {results.Failed}");
        Console.WriteLine($"Skipped: {results.Skipped}");
        Console.WriteLine("================================================================================");

        if (results.PassRate >= 0.95)
        {
            Console.WriteLine("✓ VALIDATION PASSED (≥95% pass rate)");
        }
        else
        {
            Console.WriteLine($"✗ VALIDATION FAILED (pass rate {results.PassRate:P2} is below 95% threshold)");
        }
    }

    static bool CompareValues(EvalCellValue ourValue, string excelValue, CellValues? excelDataType)
    {
        if (ourValue.IsError)
        {
            return excelValue == ourValue.ErrorValue;
        }

        if (ourValue.Type == CellValueType.Number)
        {
            if (double.TryParse(excelValue, out var excelNumber))
            {
                return Math.Abs(ourValue.NumericValue - excelNumber) < 0.0001;
            }
            return false;
        }

        if (ourValue.Type == CellValueType.Text)
        {
            return string.Equals(ourValue.StringValue, excelValue, StringComparison.Ordinal);
        }

        if (ourValue.Type == CellValueType.Boolean)
        {
            if (excelDataType == CellValues.Boolean)
            {
                return (ourValue.BoolValue && excelValue == "1") ||
                       (!ourValue.BoolValue && excelValue == "0");
            }
            return (ourValue.BoolValue && excelValue.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) ||
                   (!ourValue.BoolValue && excelValue.Equals("FALSE", StringComparison.OrdinalIgnoreCase));
        }

        return false;
    }

    static string FormatValue(EvalCellValue value)
    {
        if (value.IsError) return value.ErrorValue ?? "ERROR";
        return value.Type switch
        {
            CellValueType.Number => value.NumericValue.ToString(),
            CellValueType.Text => value.StringValue,
            CellValueType.Boolean => value.BoolValue.ToString(),
            _ => value.ToString() ?? "NULL",
        };
    }

    static string GetSheetName(WorkbookPart workbookPart, WorksheetPart worksheetPart)
    {
        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Id == workbookPart.GetIdOfPart(worksheetPart));
        return sheet?.Name?.Value ?? "Unknown";
    }

    static int GetRowNumber(string cellRef)
    {
        return int.Parse(new string(cellRef.Where(char.IsDigit).ToArray()));
    }

    class ValidationResults
    {
        public int Total { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public int Skipped { get; set; }
        public double PassRate => Total > 0 ? (double)Passed / (Total - Skipped) : 0;
    }
}
