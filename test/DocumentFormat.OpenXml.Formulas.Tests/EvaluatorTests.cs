// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Tests for end-to-end formula evaluation.
/// </summary>
public class EvaluatorTests
{
    [Fact]
    public void Evaluate_SimpleFormula_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
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
        var row = new Row();
        sheetData.Append(row);

        row.Append(new Cell
        {
            CellReference = "A1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("10"),
            DataType = CellValues.Number,
        });

        row.Append(new Cell
        {
            CellReference = "B1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("20"),
            DataType = CellValues.Number,
        });

        var cellWithFormula = new Cell
        {
            CellReference = "C1",
            CellFormula = new CellFormula("A1+B1"),
        };
        row.Append(cellWithFormula);

        document.AddFormulaEvaluationFeature();
        var evaluator = document.GetFormulaEvaluator();

        // Act
        var result = evaluator!.TryEvaluate(worksheetPart.Worksheet, cellWithFormula);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal(30.0, result.Value.NumericValue);
    }

    [Fact]
    public void Evaluate_SumFormula_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
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

        // Add cells A1-A10 with values 1-10
        for (var i = 1; i <= 10; i++)
        {
            var row = new Row();
            sheetData.Append(row);
            row.Append(new Cell
            {
                CellReference = $"A{i}",
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(i.ToString()),
                DataType = CellValues.Number,
            });
        }

        var formulaRow = new Row();
        sheetData.Append(formulaRow);

        var cellWithFormula = new Cell
        {
            CellReference = "B1",
            CellFormula = new CellFormula("SUM(A1:A10)"),
        };
        formulaRow.Append(cellWithFormula);

        document.AddFormulaEvaluationFeature();
        var evaluator = document.GetFormulaEvaluator();

        // Act
        var result = evaluator!.TryEvaluate(worksheetPart.Worksheet, cellWithFormula);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal(55.0, result.Value.NumericValue); // 1+2+...+10 = 55
    }

    [Fact]
    public void Evaluate_AverageFormula_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
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

        // Add cells A1-A5 with values 10, 20, 30, 40, 50
        for (var i = 1; i <= 5; i++)
        {
            var row = new Row();
            sheetData.Append(row);
            row.Append(new Cell
            {
                CellReference = $"A{i}",
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue((i * 10).ToString()),
                DataType = CellValues.Number,
            });
        }

        var formulaRow = new Row();
        sheetData.Append(formulaRow);

        var cellWithFormula = new Cell
        {
            CellReference = "B1",
            CellFormula = new CellFormula("AVERAGE(A1:A5)"),
        };
        formulaRow.Append(cellWithFormula);

        document.AddFormulaEvaluationFeature();
        var evaluator = document.GetFormulaEvaluator();

        // Act
        var result = evaluator!.TryEvaluate(worksheetPart.Worksheet, cellWithFormula);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal(30.0, result.Value.NumericValue); // (10+20+30+40+50)/5 = 30
    }

    [Fact]
    public void Evaluate_IfFormula_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
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
        var row = new Row();
        sheetData.Append(row);

        row.Append(new Cell
        {
            CellReference = "A1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("15"),
            DataType = CellValues.Number,
        });

        row.Append(new Cell
        {
            CellReference = "B1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("100"),
            DataType = CellValues.Number,
        });

        row.Append(new Cell
        {
            CellReference = "C1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("50"),
            DataType = CellValues.Number,
        });

        var cellWithFormula = new Cell
        {
            CellReference = "D1",
            CellFormula = new CellFormula("IF(A1>10, B1, C1)"),
        };
        row.Append(cellWithFormula);

        document.AddFormulaEvaluationFeature();
        var evaluator = document.GetFormulaEvaluator();

        // Act
        var result = evaluator!.TryEvaluate(worksheetPart.Worksheet, cellWithFormula);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.Equal(100.0, result.Value.NumericValue); // A1 (15) > 10, so returns B1 (100)
    }

    [Fact]
    public void Evaluate_CachesBehavior_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
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
        var row = new Row();
        sheetData.Append(row);

        row.Append(new Cell
        {
            CellReference = "A1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("10"),
            DataType = CellValues.Number,
        });

        row.Append(new Cell
        {
            CellReference = "B1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("20"),
            DataType = CellValues.Number,
        });

        var cellWithFormula = new Cell
        {
            CellReference = "C1",
            CellFormula = new CellFormula("A1+B1"),
        };
        row.Append(cellWithFormula);

        document.AddFormulaEvaluationFeature();
        var evaluator = document.GetFormulaEvaluator();

        // Act - evaluate twice
        var result1 = evaluator!.TryEvaluate(worksheetPart.Worksheet, cellWithFormula);
        var result2 = evaluator.TryEvaluate(worksheetPart.Worksheet, cellWithFormula);

        // Assert - both should return the same result
        Assert.True(result1.IsSuccess);
        Assert.True(result2.IsSuccess);
        Assert.Equal(result1.Value.NumericValue, result2.Value.NumericValue);
    }

    [Fact(Skip = "Phase 0 limitation - cache does not invalidate")]
    public void Evaluate_CacheDoesNotInvalidate_KnownLimitation()
    {
        // This test demonstrates the known Phase 0 limitation that the cell value cache
        // does not invalidate when cell values change. This will be fixed in Phase 1.

        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // Add cells: A1=10, A2=A1*2 (formula)
            var row1 = new Row { RowIndex = 1 };
            var cellA1 = new Cell
            {
                CellReference = "A1",
                DataType = CellValues.Number,
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("10"),
            };
            row1.Append(cellA1);
            sheetData.Append(row1);

            var row2 = new Row { RowIndex = 2 };
            var cellA2 = new Cell
            {
                CellReference = "A2",
                CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula("A1*2"),
            };
            row2.Append(cellA2);
            sheetData.Append(row2);

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // First evaluation: A2 = A1*2 = 10*2 = 20
            var result1 = evaluator.TryEvaluate(wsPart.Worksheet, cellA2);
            Assert.True(result1.IsSuccess);
            Assert.Equal(20.0, result1.Value.NumericValue);

            // Change A1 value
            cellA1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("5");

            // Second evaluation: SHOULD be A2 = A1*2 = 5*2 = 10
            // But due to caching limitation, will return 20 (stale value)
            var result2 = evaluator.TryEvaluate(wsPart.Worksheet, cellA2);
            Assert.True(result2.IsSuccess);

            // This assertion will fail in Phase 0 (returns 20), succeed in Phase 1 (returns 10)
            Assert.Equal(10.0, result2.Value.NumericValue);
        }
    }

    [Fact]
    public void Evaluate_RankFormula_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
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

        // Add cells F1-F10 with values 5, 10, 15, 20, 25, 30, 35, 40, 45, 50
        for (var i = 1; i <= 10; i++)
        {
            var row = new Row { RowIndex = (uint)i };
            sheetData.Append(row);
            row.Append(new Cell
            {
                CellReference = $"F{i}",
                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue((i * 5).ToString()),
                DataType = CellValues.Number,
            });
        }

        var formulaRow = new Row { RowIndex = 11 };
        sheetData.Append(formulaRow);

        var cellWithFormula = new Cell
        {
            CellReference = "A11",
            CellFormula = new CellFormula("RANK(25, F1:F10)"),
        };
        formulaRow.Append(cellWithFormula);

        document.AddFormulaEvaluationFeature();
        var evaluator = document.GetFormulaEvaluator();

        // Act
        var result = evaluator!.TryEvaluate(worksheetPart.Worksheet, cellWithFormula);

        // Assert
        Assert.True(result.IsSuccess);
        // 25 is the 5th largest value in [5, 10, 15, 20, 25, 30, 35, 40, 45, 50]
        // In descending order: 50(1), 45(2), 40(3), 35(4), 30(5), 25(6), 20(7), ...
        Assert.Equal(6.0, result.Value.NumericValue);
    }
}
