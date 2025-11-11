// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;

using BenchmarkDotNet.Attributes;

using DocumentFormat.OpenXml.Features.FormulaEvaluation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Benchmarks;

/// <summary>
/// Benchmarks for formula evaluation performance.
/// </summary>
[MemoryDiagnoser]
public class FormulaEvaluationBenchmarks
{
    private SpreadsheetDocument? _document;
    private Cell? _simpleFormulaCell;
    private Cell? _sumFormulaCell;
    private IFormulaEvaluator? _evaluator;

    /// <summary>
    /// Sets up the benchmark.
    /// </summary>
    [GlobalSetup]
    public void Setup()
    {
        var stream = new MemoryStream();
        _document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = _document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheet = new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1",
        };
        sheets.Append(sheet);

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

        // Add data cells
        var row1 = new Row();
        sheetData.Append(row1);

        row1.Append(new Cell
        {
            CellReference = "A1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("10"),
            DataType = CellValues.Number,
        });

        row1.Append(new Cell
        {
            CellReference = "B1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("20"),
            DataType = CellValues.Number,
        });

        // Add simple formula cell
        _simpleFormulaCell = new Cell
        {
            CellReference = "C1",
            CellFormula = new CellFormula("A1+B1"),
        };
        row1.Append(_simpleFormulaCell);

        // Add cells for SUM benchmark
        for (var i = 1; i <= 100; i++)
        {
            var row = new Row();
            sheetData.Append(row);
            row.Append(new Cell
            {
                CellReference = $"D{i}",
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(i.ToString()),
                DataType = CellValues.Number,
            });
        }

        // Add SUM formula cell
        var sumRow = new Row();
        sheetData.Append(sumRow);
        _sumFormulaCell = new Cell
        {
            CellReference = "E1",
            CellFormula = new CellFormula("SUM(D1:D100)"),
        };
        sumRow.Append(_sumFormulaCell);

        _document.AddFormulaEvaluationFeature();
        _evaluator = _document.GetFormulaEvaluator();
    }

    /// <summary>
    /// Cleans up after benchmarks.
    /// </summary>
    [GlobalCleanup]
    public void Cleanup()
    {
        _evaluator?.Dispose();
        _document?.Dispose();
    }

    /// <summary>
    /// Benchmarks evaluating a simple formula (A1+B1).
    /// </summary>
    [Benchmark]
    public void EvaluateSimpleFormula()
    {
        _evaluator!.TryEvaluate(_simpleFormulaCell!);
    }

    /// <summary>
    /// Benchmarks evaluating a SUM formula over 100 cells.
    /// </summary>
    [Benchmark]
    public void EvaluateSumFormula()
    {
        _evaluator!.TryEvaluate(_sumFormulaCell!);
    }
}
