// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation.DependencyGraph;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using Xunit;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Tests for dependency graph implementation and Kahn's algorithm.
/// Tests the public API through IFormulaEvaluator.GetDependencyGraph().
/// </summary>
public class DependencyGraphTests
{
    [Fact]
    public void GetEvaluationOrder_SimpleChain_ReturnsCorrectOrder()
    {
        // A1 = 10
        // A2 = A1 * 2
        // A3 = A2 + 5
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
            AddFormulaCell(sheetData, "A3", "A2+5");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var order = graph.GetEvaluationOrder();

            // A2 must come before A3
            Assert.Equal(2, order.Count);
            Assert.Equal("A2", order[0]);
            Assert.Equal("A3", order[1]);
        }
    }

    [Fact]
    public void GetEvaluationOrder_IndependentCells_AnyOrderValid()
    {
        // A1 = 10
        // B1 = 20
        // C1 = A1 + B1
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
            AddCell(sheetData, "B1", "20");
            AddFormulaCell(sheetData, "C1", "A1+B1");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var order = graph.GetEvaluationOrder();

            Assert.Single(order);
            Assert.Equal("C1", order[0]);
        }
    }

    [Fact]
    public void GetEvaluationOrder_DiamondDependency_ReturnsValidOrder()
    {
        // A1 = 10
        // B1 = A1 * 2
        // C1 = A1 * 3
        // D1 = B1 + C1
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
            AddFormulaCell(sheetData, "B1", "A1*2");
            AddFormulaCell(sheetData, "C1", "A1*3");
            AddFormulaCell(sheetData, "D1", "B1+C1");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var order = graph.GetEvaluationOrder();

            Assert.Equal(3, order.Count);
            var d1Index = order.IndexOf("D1");
            var b1Index = order.IndexOf("B1");
            var c1Index = order.IndexOf("C1");

            // D1 must come after both B1 and C1
            Assert.True(d1Index > b1Index);
            Assert.True(d1Index > c1Index);
        }
    }

    [Fact]
    public void GetEvaluationOrder_ComplexDependencies_ReturnsValidOrder()
    {
        // More complex dependency graph:
        // A1 = 5
        // A2 = A1 + 1
        // A3 = A1 + A2
        // A4 = A2 * 2
        // A5 = A3 + A4
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

            AddCell(sheetData, "A1", "5");
            AddFormulaCell(sheetData, "A2", "A1+1");
            AddFormulaCell(sheetData, "A3", "A1+A2");
            AddFormulaCell(sheetData, "A4", "A2*2");
            AddFormulaCell(sheetData, "A5", "A3+A4");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var order = graph.GetEvaluationOrder();

            Assert.Equal(4, order.Count);

            // Verify topological constraints
            var a2Index = order.IndexOf("A2");
            var a3Index = order.IndexOf("A3");
            var a4Index = order.IndexOf("A4");
            var a5Index = order.IndexOf("A5");

            // A2 must come before A3, A4, and A5
            Assert.True(a2Index < a3Index);
            Assert.True(a2Index < a4Index);
            Assert.True(a2Index < a5Index);

            // A3 and A4 must come before A5
            Assert.True(a3Index < a5Index);
            Assert.True(a4Index < a5Index);
        }
    }

    [Fact]
    public void DetectCircularReferences_SimpleCircle_DetectsCorrectly()
    {
        // A1 = A2
        // A2 = A1  (circular!)
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

            AddFormulaCell(sheetData, "A1", "A2");
            AddFormulaCell(sheetData, "A2", "A1");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var cycles = graph.DetectCircularReferences();

            Assert.Single(cycles);
            Assert.Contains("A1", cycles[0].Chain);
            Assert.Contains("A2", cycles[0].Chain);
        }
    }

    [Fact]
    public void DetectCircularReferences_LongerCircle_DetectsCorrectly()
    {
        // A1 = A2
        // A2 = A3
        // A3 = A1  (circular!)
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

            AddFormulaCell(sheetData, "A1", "A2");
            AddFormulaCell(sheetData, "A2", "A3");
            AddFormulaCell(sheetData, "A3", "A1");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var cycles = graph.DetectCircularReferences();

            Assert.Single(cycles);
            Assert.Equal(4, cycles[0].Chain.Count); // A1 → A2 → A3 → A1
        }
    }

    [Fact]
    public void DetectCircularReferences_NoCircle_ReturnsEmpty()
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
            AddFormulaCell(sheetData, "A3", "A2+5");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var cycles = graph.DetectCircularReferences();

            Assert.Empty(cycles);
        }
    }

    [Fact]
    public void GetDependencies_ExistingCell_ReturnsCorrectDependencies()
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
            AddCell(sheetData, "A2", "20");
            AddFormulaCell(sheetData, "A3", "A1+A2");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var deps = graph.GetDependencies("A3");

            Assert.Equal(2, deps.Count);
            Assert.Contains("A1", deps);
            Assert.Contains("A2", deps);
        }
    }

    [Fact]
    public void GetDependencies_NonExistingCell_ReturnsEmptySet()
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

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var deps = graph.GetDependencies("A1");

            Assert.Empty(deps);
        }
    }

    [Fact]
    public void GetDependents_ExistingCell_ReturnsCorrectDependents()
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
            AddFormulaCell(sheetData, "A3", "A1+5");

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var dependents = graph.GetDependents("A1");

            Assert.Equal(2, dependents.Count);
            Assert.Contains("A2", dependents);
            Assert.Contains("A3", dependents);
        }
    }

    [Fact]
    public void GetDependents_NonExistingCell_ReturnsEmptySet()
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

            doc.AddFormulaEvaluationFeature();
            var evaluator = doc.Features.GetRequired<IFormulaEvaluator>();

            var graph = evaluator.GetDependencyGraph(wsPart.Worksheet);
            var dependents = graph.GetDependents("A1");

            Assert.Empty(dependents);
        }
    }

    [Fact]
    public void CircularReference_ToString_ReturnsFormattedChain()
    {
        var chain = new List<string> { "A1", "A2", "A3", "A1" };
        var circular = new CircularReference(chain);

        var result = circular.ToString();

        Assert.Equal("A1 → A2 → A3 → A1", result);
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
