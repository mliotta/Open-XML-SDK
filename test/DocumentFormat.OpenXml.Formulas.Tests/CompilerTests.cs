// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Parsing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Tests for formula compilation.
/// </summary>
public class CompilerTests
{
    [Fact]
    public void Compile_SimpleAddition_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

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

        var parser = new FormulaParser();
        var compiler = new FormulaCompiler();
        var context = new CellContext(worksheetPart.Worksheet);

        // Act
        var ast = parser.Parse("=A1+B1");
        var expression = compiler.Compile(ast);
        var compiled = expression.Compile();
        var result = compiled(context);

        // Assert
        Assert.Equal(CellValueType.Number, result.Type);
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Compile_Multiplication_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        var row = new Row();
        sheetData.Append(row);

        row.Append(new Cell
        {
            CellReference = "A1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("5"),
            DataType = CellValues.Number,
        });

        row.Append(new Cell
        {
            CellReference = "B1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("6"),
            DataType = CellValues.Number,
        });

        var parser = new FormulaParser();
        var compiler = new FormulaCompiler();
        var context = new CellContext(worksheetPart.Worksheet);

        // Act
        var ast = parser.Parse("=A1*B1");
        var expression = compiler.Compile(ast);
        var compiled = expression.Compile();
        var result = compiled(context);

        // Assert
        Assert.Equal(30.0, result.NumericValue);
    }

    [Fact]
    public void Compile_OperatorPrecedence_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        var row = new Row();
        sheetData.Append(row);

        row.Append(new Cell
        {
            CellReference = "A1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("2"),
            DataType = CellValues.Number,
        });

        row.Append(new Cell
        {
            CellReference = "B1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("3"),
            DataType = CellValues.Number,
        });

        row.Append(new Cell
        {
            CellReference = "C1",
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("4"),
            DataType = CellValues.Number,
        });

        var parser = new FormulaParser();
        var compiler = new FormulaCompiler();
        var context = new CellContext(worksheetPart.Worksheet);

        // Act
        var ast = parser.Parse("=A1+B1*C1"); // Should be 2 + (3 * 4) = 14
        var expression = compiler.Compile(ast);
        var compiled = expression.Compile();
        var result = compiled(context);

        // Assert
        Assert.Equal(14.0, result.NumericValue);
    }

    [Fact]
    public void Compile_SumFunction_Success()
    {
        // Arrange
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

        // Add cells A1-A5 with values 1-5
        for (var i = 1; i <= 5; i++)
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

        var parser = new FormulaParser();
        var compiler = new FormulaCompiler();
        var context = new CellContext(worksheetPart.Worksheet);

        // Act
        var ast = parser.Parse("=SUM(A1:A5)");
        var expression = compiler.Compile(ast);
        var compiled = expression.Compile();
        var result = compiled(context);

        // Assert
        Assert.Equal(15.0, result.NumericValue); // 1+2+3+4+5 = 15
    }
}
