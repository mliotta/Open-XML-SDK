// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Oracle validation tests that compare our formula evaluation results
/// against Excel's cached calculated values.
/// </summary>
public class OracleValidationTests
{

    /// <summary>
    /// Generates the oracle test file. Run this test once to create the file,
    /// then open it in Excel to bless the formulas with Excel's calculations.
    /// </summary>
    [Fact]
    public void GenerateOracleTestFile()
    {
        var filePath = Path.Combine(Path.GetTempPath(), "FormulaOracle.xlsx");
        OracleTestFileGenerator.GenerateOracleTestFile(filePath);

        Console.WriteLine($"Oracle test file generated at: {filePath}");
        Console.WriteLine(string.Empty);
        Console.WriteLine("NEXT STEPS:");
        Console.WriteLine("1. Open this file in Excel");
        Console.WriteLine("2. Excel will calculate all formulas and store cached values");
        Console.WriteLine("3. Save and close the file");
        Console.WriteLine("4. Copy the file to: test/DocumentFormat.OpenXml.Formulas.Tests/TestFiles/FormulaOracle.xlsx");
        Console.WriteLine("5. Unskip the ValidateAgainstExcelOracle test");
    }

    /// <summary>
    /// Validates our formula evaluation against Excel's cached results.
    /// This test requires the oracle file to be blessed by Excel first.
    /// </summary>
    [Fact]
    public void ValidateAgainstExcelOracle()
    {
        var oracleFilePath = Path.Combine(
            Directory.GetCurrentDirectory(),
            "TestFiles",
            "FormulaOracle.xlsx");

        if (!File.Exists(oracleFilePath))
        {
            throw new FileNotFoundException(
                "Oracle file not found. Run GenerateOracleTestFile first and bless it with Excel.",
                oracleFilePath);
        }

        var results = ValidateOracleFile(oracleFilePath);

        // Report results
        Console.WriteLine($"Total test cases: {results.Total}");
        Console.WriteLine($"Passed: {results.Passed} ({results.PassRate:P2})");
        Console.WriteLine($"Failed: {results.Failed}");
        Console.WriteLine($"Skipped: {results.Skipped}");
        Console.WriteLine(string.Empty);

        if (results.Failures.Count > 0)
        {
            Console.WriteLine("FAILURES:");
            foreach (var failure in results.Failures)
            {
                Console.WriteLine($"  {failure.Sheet}!{failure.CellRef}: {failure.Formula}");
                Console.WriteLine($"    Expected: {failure.ExpectedValue}");
                Console.WriteLine($"    Actual:   {failure.ActualValue}");
                Console.WriteLine($"    Error:    {failure.ErrorMessage}");
                Console.WriteLine(string.Empty);
            }
        }

        // Require 95% pass rate
        Assert.True(results.PassRate >= 0.95,
            $"Pass rate {results.PassRate:P2} is below 95% threshold. {results.Failed} failures out of {results.Total} tests.");
    }

    private OracleValidationResult ValidateOracleFile(string filePath)
    {
        var result = new OracleValidationResult();

        using var doc = SpreadsheetDocument.Open(filePath, false);
        doc.AddFormulaEvaluationFeature();
        var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

        foreach (var worksheetPart in doc.WorkbookPart!.WorksheetParts)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetName = GetSheetName(doc.WorkbookPart, worksheetPart);

            // Column C contains formulas with Excel-calculated cached values
            var formulaCells = worksheet.Descendants<Cell>()
                .Where(c => c.CellReference != null &&
                           c.CellReference.Value!.StartsWith("C") &&
                           c.CellFormula != null)
                .OrderBy(c => GetRowNumber(c.CellReference!.Value!));

            foreach (var cell in formulaCells)
            {
                result.Total++;

                var cellRef = cell.CellReference!.Value!;
                var formula = cell.CellFormula!.Text;

                // Skip header row
                if (GetRowNumber(cellRef) == 1)
                {
                    result.Skipped++;
                    continue;
                }

                // Get Excel's cached value
                var excelValue = cell.CellValue?.Text;
                var excelDataType = cell.DataType?.Value;

                if (string.IsNullOrEmpty(excelValue))
                {
                    result.Skipped++;
                    Console.WriteLine($"Skipped {sheetName}!{cellRef}: No Excel cached value");
                    continue;
                }

                // Evaluate with our engine
                var evalResult = evaluator.TryEvaluate(worksheet, cell);

                if (!evalResult.IsSuccess)
                {
                    result.Failed++;
                    result.Failures.Add(new OracleFailure
                    {
                        Sheet = sheetName,
                        CellRef = cellRef,
                        Formula = formula,
                        ExpectedValue = excelValue,
                        ActualValue = null,
                        ErrorMessage = evalResult.Error?.Message ?? "Unknown error",
                    });
                    continue;
                }

                // Compare values
                var ourValue = evalResult.Value;
                var match = CompareValues(ourValue, excelValue, excelDataType);

                if (match)
                {
                    result.Passed++;
                }
                else
                {
                    result.Failed++;
                    result.Failures.Add(new OracleFailure
                    {
                        Sheet = sheetName,
                        CellRef = cellRef,
                        Formula = formula,
                        ExpectedValue = excelValue,
                        ActualValue = FormatValue(ourValue),
                        ErrorMessage = "Value mismatch",
                    });
                }
            }
        }

        return result;
    }

    private bool CompareValues(CellValue ourValue, string excelValue, CellValues? excelDataType)
    {
        // Handle errors
        if (ourValue.IsError)
        {
            return excelValue == ourValue.ErrorValue;
        }

        // Handle numbers
        if (ourValue.Type == CellValueType.Number)
        {
            if (double.TryParse(excelValue, out var excelNumber))
            {
                // Allow small floating point differences
                return System.Math.Abs(ourValue.NumericValue - excelNumber) < 0.0001;
            }

            return false;
        }

        // Handle text
        if (ourValue.Type == CellValueType.Text)
        {
            return string.Equals(ourValue.StringValue, excelValue, StringComparison.Ordinal);
        }

        // Handle booleans
        if (ourValue.Type == CellValueType.Boolean)
        {
            // Excel stores TRUE as "1" and FALSE as "0"
            if (excelDataType == CellValues.Boolean)
            {
                return (ourValue.BoolValue && excelValue == "1") ||
                       (!ourValue.BoolValue && excelValue == "0");
            }

            // Or as text
            return (ourValue.BoolValue && excelValue.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) ||
                   (!ourValue.BoolValue && excelValue.Equals("FALSE", StringComparison.OrdinalIgnoreCase));
        }

        return false;
    }

    private string FormatValue(CellValue value)
    {
        if (value.IsError)
        {
            return value.ErrorValue ?? "ERROR";
        }

        return value.Type switch
        {
            CellValueType.Number => value.NumericValue.ToString(),
            CellValueType.Text => value.StringValue,
            CellValueType.Boolean => value.BoolValue.ToString(),
            _ => value.ToString() ?? "NULL",
        };
    }

    private string GetSheetName(WorkbookPart workbookPart, WorksheetPart worksheetPart)
    {
        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Id == workbookPart.GetIdOfPart(worksheetPart));

        return sheet?.Name?.Value ?? "Unknown";
    }

    private int GetRowNumber(string cellRef)
    {
        return int.Parse(new string(cellRef.Where(char.IsDigit).ToArray()));
    }

    private class OracleValidationResult
    {
        public int Total { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public int Skipped { get; set; }
        public List<OracleFailure> Failures { get; } = new();

        public double PassRate => Total > 0 ? (double)Passed / (Total - Skipped) : 0;
    }

    private class OracleFailure
    {
        public string Sheet { get; set; } = string.Empty;
        public string CellRef { get; set; } = string.Empty;
        public string Formula { get; set; } = string.Empty;
        public string? ExpectedValue { get; set; }
        public string? ActualValue { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
