// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Tests for incremental recalculation (RecalculateDependents).
/// </summary>
public class RecalculateDependentsTests
{
    [Fact]
    public void RecalculateDependents_SingleDependent_UpdatesOnlyThatCell()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // A1=10, A2=A1*2, A3=5
            AddCell(sheetData, "A1", "10");
            AddFormulaCell(sheetData, "A2", "A1*2");
            AddCell(sheetData, "A3", "5");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Initial calculation
            evaluator.RecalculateSheet(wsPart.Worksheet);

            // Change A1
            var a1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
            a1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("20");

            // Recalculate only dependents of A1
            evaluator.RecalculateDependents(wsPart.Worksheet, "A1");

            // Verify A2 was updated
            var a2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
            Assert.NotNull(a2.CellValue);
            Assert.Equal("40", a2.CellValue.Text);
        }
    }

    [Fact]
    public void RecalculateDependents_DependencyChain_UpdatesAllInChain()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // A1=10, A2=A1*2, A3=A2+5, A4=A3-3
            AddCell(sheetData, "A1", "10");
            AddFormulaCell(sheetData, "A2", "A1*2");
            AddFormulaCell(sheetData, "A3", "A2+5");
            AddFormulaCell(sheetData, "A4", "A3-3");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Initial calculation
            evaluator.RecalculateSheet(wsPart.Worksheet);

            // Change A1
            var a1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
            a1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("20");

            // Recalculate dependents
            evaluator.RecalculateDependents(wsPart.Worksheet, "A1");

            // Verify entire chain updated
            var a2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
            Assert.NotNull(a2.CellValue);
            Assert.Equal("40", a2.CellValue.Text); // 20*2

            var a3 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A3");
            Assert.NotNull(a3.CellValue);
            Assert.Equal("45", a3.CellValue.Text); // 40+5

            var a4 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A4");
            Assert.NotNull(a4.CellValue);
            Assert.Equal("42", a4.CellValue.Text); // 45-3
        }
    }

    [Fact]
    public void RecalculateDependents_MultipleSources_UpdatesAllDependents()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // A1=10, B1=20, C1=A1+B1
            AddCell(sheetData, "A1", "10");
            AddCell(sheetData, "B1", "20");
            AddFormulaCell(sheetData, "C1", "A1+B1");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Initial calculation
            evaluator.RecalculateSheet(wsPart.Worksheet);

            // Change both A1 and B1
            var a1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
            a1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("15");

            var b1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
            b1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("25");

            // Recalculate dependents of both
            evaluator.RecalculateDependents(wsPart.Worksheet, "A1", "B1");

            // Verify C1 was updated
            var c1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C1");
            Assert.NotNull(c1.CellValue);
            Assert.Equal("40", c1.CellValue.Text); // 15+25
        }
    }

    [Fact]
    public void RecalculateDependents_DiamondPattern_EvaluatesInCorrectOrder()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // A1=10, B1=A1*2, C1=A1*3, D1=B1+C1 (diamond pattern)
            AddCell(sheetData, "A1", "10");
            AddFormulaCell(sheetData, "B1", "A1*2");
            AddFormulaCell(sheetData, "C1", "A1*3");
            AddFormulaCell(sheetData, "D1", "B1+C1");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Initial calculation
            evaluator.RecalculateSheet(wsPart.Worksheet);

            // Change A1
            var a1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
            a1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("20");

            // Recalculate dependents
            evaluator.RecalculateDependents(wsPart.Worksheet, "A1");

            // Verify all dependents updated correctly
            var b1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
            Assert.NotNull(b1.CellValue);
            Assert.Equal("40", b1.CellValue.Text); // 20*2

            var c1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C1");
            Assert.NotNull(c1.CellValue);
            Assert.Equal("60", c1.CellValue.Text); // 20*3

            var d1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D1");
            Assert.NotNull(d1.CellValue);
            Assert.Equal("100", d1.CellValue.Text); // 40+60
        }
    }

    [Fact]
    public void RecalculateDependents_NoDependents_DoesNothing()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // A1=10, A2=20 (no dependencies)
            AddCell(sheetData, "A1", "10");
            AddCell(sheetData, "A2", "20");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Change A1 (has no dependents)
            var a1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
            a1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("50");

            // Should not throw
            evaluator.RecalculateDependents(wsPart.Worksheet, "A1");

            // A2 should be unchanged
            var a2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
            Assert.NotNull(a2.CellValue);
            Assert.Equal("20", a2.CellValue.Text);
        }
    }

    [Fact]
    public void RecalculateDependents_EmptyArray_DoesNothing()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Should not throw
            evaluator.RecalculateDependents(wsPart.Worksheet);
        }
    }

    [Fact]
    public void RecalculateDependents_RangeDependent_UpdatesCorrectly()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // A1=10, A2=20, A3=30, B1=SUM(A1:A3)
            AddCell(sheetData, "A1", "10");
            AddCell(sheetData, "A2", "20");
            AddCell(sheetData, "A3", "30");
            AddFormulaCell(sheetData, "B1", "SUM(A1:A3)");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Initial calculation
            evaluator.RecalculateSheet(wsPart.Worksheet);

            // Change A2 (part of range)
            var a2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
            a2.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("50");

            // Recalculate dependents
            evaluator.RecalculateDependents(wsPart.Worksheet, "A2");

            // Verify B1 was updated
            var b1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
            Assert.NotNull(b1.CellValue);
            Assert.Equal("90", b1.CellValue.Text); // 10+50+30
        }
    }

    [Fact]
    public void RecalculateDependents_PartialUpdate_DoesNotRecalculateIndependentCells()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            wbPart.Workbook.AppendChild(new Sheets()).Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Sheet1",
            });

            // A1=10, B1=20, A2=A1*2, B2=B1*2
            AddCell(sheetData, "A1", "10");
            AddCell(sheetData, "B1", "20");
            AddFormulaCell(sheetData, "A2", "A1*2");
            AddFormulaCell(sheetData, "B2", "B1*2");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Initial calculation
            evaluator.RecalculateSheet(wsPart.Worksheet);

            // Change A1 only
            var a1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
            a1.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("30");

            // Recalculate dependents (should only update A2, not B2)
            evaluator.RecalculateDependents(wsPart.Worksheet, "A1");

            // Verify A2 was updated
            var a2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A2");
            Assert.NotNull(a2.CellValue);
            Assert.Equal("60", a2.CellValue.Text); // 30*2

            // B2 should still have old value from initial calculation
            var b2 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B2");
            Assert.NotNull(b2.CellValue);
            Assert.Equal("40", b2.CellValue.Text); // 20*2 (unchanged)
        }
    }

    private static void AddCell(SheetData sheetData, string reference, string value)
    {
        var rowIndex = GetRowIndex(reference);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        row.Append(new Cell
        {
            CellReference = reference,
            DataType = CellValues.Number,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value),
        });
    }

    private static void AddFormulaCell(SheetData sheetData, string reference, string formula)
    {
        var rowIndex = GetRowIndex(reference);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        row.Append(new Cell
        {
            CellReference = reference,
            CellFormula = new CellFormula(formula),
        });
    }

    private static uint GetRowIndex(string reference)
    {
        return uint.Parse(new string(reference.Where(char.IsDigit).ToArray()));
    }
}
