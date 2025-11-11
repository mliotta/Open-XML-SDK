// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Integration tests for RecalculateSheet and end-to-end formula evaluation.
/// </summary>
public class IntegrationTests
{
    [Fact]
    public void RecalculateSheet_SimpleWorkbook_UpdatesAllValues()
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

            // Add: A1=10, A2=20, A3=A1+A2
            AddCell(sheetData, "A1", "10");
            AddCell(sheetData, "A2", "20");
            AddFormulaCell(sheetData, "A3", "A1+A2");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            evaluator.RecalculateSheet(wsPart.Worksheet);

            var cellA3 = wsPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference == "A3");

            Assert.NotNull(cellA3.CellValue);
            Assert.Equal("30", cellA3.CellValue.Text);
        }
    }

    [Fact]
    public void RecalculateSheet_DependencyChain_EvaluatesInCorrectOrder()
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

            // A1=5, A2=A1*2, A3=A2+10, A4=A3*3
            AddCell(sheetData, "A1", "5");
            AddFormulaCell(sheetData, "A2", "A1*2");
            AddFormulaCell(sheetData, "A3", "A2+10");
            AddFormulaCell(sheetData, "A4", "A3*3");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            evaluator.RecalculateSheet(wsPart.Worksheet);

            var cellA4 = wsPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference == "A4");

            // A2=10, A3=20, A4=60
            Assert.NotNull(cellA4.CellValue);
            Assert.Equal("60", cellA4.CellValue.Text);
        }
    }

    [Fact]
    public void RecalculateSheet_ComplexDependencies_EvaluatesCorrectly()
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

            // Diamond dependency:
            // A1=10, B1=A1*2, C1=A1*3, D1=B1+C1
            AddCell(sheetData, "A1", "10");
            AddFormulaCell(sheetData, "B1", "A1*2");
            AddFormulaCell(sheetData, "C1", "A1*3");
            AddFormulaCell(sheetData, "D1", "B1+C1");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            evaluator.RecalculateSheet(wsPart.Worksheet);

            var cellD1 = wsPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference == "D1");

            // B1=20, C1=30, D1=50
            Assert.NotNull(cellD1.CellValue);
            Assert.Equal("50", cellD1.CellValue.Text);
        }
    }

    [Fact]
    public void RecalculateSheet_WithFunctions_EvaluatesCorrectly()
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

            // Create range A1:A5 with values 10, 20, 30, 40, 50
            for (int i = 1; i <= 5; i++)
            {
                AddCell(sheetData, $"A{i}", (i * 10).ToString());
            }

            // B1 = SUM(A1:A5)
            // B2 = AVERAGE(A1:A5)
            // B3 = MAX(A1:A5)
            AddFormulaCell(sheetData, "B1", "SUM(A1:A5)");
            AddFormulaCell(sheetData, "B2", "AVERAGE(A1:A5)");
            AddFormulaCell(sheetData, "B3", "MAX(A1:A5)");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            evaluator.RecalculateSheet(wsPart.Worksheet);

            var cellB1 = wsPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference == "B1");
            var cellB2 = wsPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference == "B2");
            var cellB3 = wsPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference == "B3");

            // B1 = 10+20+30+40+50 = 150
            Assert.NotNull(cellB1.CellValue);
            Assert.Equal("150", cellB1.CellValue.Text);

            // B2 = 150/5 = 30
            Assert.NotNull(cellB2.CellValue);
            Assert.Equal("30", cellB2.CellValue.Text);

            // B3 = 50
            Assert.NotNull(cellB3.CellValue);
            Assert.Equal("50", cellB3.CellValue.Text);
        }
    }

    [Fact]
    public void GetDependencyGraph_ReturnsValidGraph()
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

            // A1=10, A2=A1*2, A3=A2+5
            AddCell(sheetData, "A1", "10");
            AddFormulaCell(sheetData, "A2", "A1*2");
            AddFormulaCell(sheetData, "A3", "A2+5");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);

            Assert.NotNull(graph);

            // Check dependencies
            var a2Deps = graph.GetDependencies("A2");
            Assert.Single(a2Deps);
            Assert.Contains("A1", a2Deps);

            var a3Deps = graph.GetDependencies("A3");
            Assert.Single(a3Deps);
            Assert.Contains("A2", a3Deps);

            // Check dependents
            var a1Dependents = graph.GetDependents("A1");
            Assert.Single(a1Dependents);
            Assert.Contains("A2", a1Dependents);

            var a2Dependents = graph.GetDependents("A2");
            Assert.Single(a2Dependents);
            Assert.Contains("A3", a2Dependents);
        }
    }

    [Fact]
    public void GetStatistics_ReturnsValidStatistics()
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

            AddCell(sheetData, "A1", "10");
            AddFormulaCell(sheetData, "A2", "A1*2");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            // Evaluate once
            var cell = wsPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference == "A2");
            evaluator.TryEvaluate(wsPart.Worksheet, cell);

            var stats = evaluator.GetStatistics();

            Assert.NotNull(stats);
            Assert.True(stats.TotalEvaluations > 0);
            Assert.True(stats.SuccessfulEvaluations > 0);
            Assert.True(stats.SuccessRate >= 0 && stats.SuccessRate <= 1.0);
        }
    }

    [Fact]
    public void IsFunctionSupported_KnownFunction_ReturnsTrue()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            Assert.True(evaluator.IsFunctionSupported("SUM"));
            Assert.True(evaluator.IsFunctionSupported("AVERAGE"));
            Assert.True(evaluator.IsFunctionSupported("IF"));
        }
    }

    [Fact]
    public void IsFunctionSupported_UnknownFunction_ReturnsFalse()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            Assert.False(evaluator.IsFunctionSupported("UNKNOWNFUNCTION"));
        }
    }

    [Fact]
    public void SupportedFunctions_ReturnsNonEmptySet()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var supported = evaluator.SupportedFunctions;

            Assert.NotEmpty(supported);
            Assert.Contains("SUM", supported);
            Assert.Contains("AVERAGE", supported);
            Assert.Contains("COUNT", supported);
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
