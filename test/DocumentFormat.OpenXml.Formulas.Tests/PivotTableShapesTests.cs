// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

using Xunit;

namespace DocumentFormat.OpenXml.Features.PivotTables.Tests;

/// <summary>
/// Covers the generalized pivot shapes: row-only, column-only, multi-row-field, and
/// multi-value-field layouts. Each asserts precomputed values and schema validity.
/// </summary>
public class PivotTableShapesTests
{
    // Region | Channel | Quarter | Sales | Units
    // East   | Online  | Q1      | 100   | 5
    // East   | Online  | Q2      | 200   | 8
    // East   | Retail  | Q1      | 10    | 1
    // West   | Online  | Q1      | 50    | 3
    // West   | Retail  | Q2      | 75    | 4
    private const string Range = "A1:E6";

    [Fact]
    public void RowOnly_SumsPerRowMemberAndGrandTotal()
    {
        using var document = CreateWorkbook(out var source, out var pivot);

        PivotTableBuilder.FromRange(source, Range).Row("Region").Value("Sales").PlaceAt(pivot, "A1");

        Assert.Equal("Sum of Sales", Text(pivot, "A1"));
        Assert.Equal("East", Text(pivot, "A2"));
        Assert.Equal("310", Text(pivot, "B2"));
        Assert.Equal("West", Text(pivot, "A3"));
        Assert.Equal("125", Text(pivot, "B3"));
        Assert.Equal("Grand Total", Text(pivot, "A4"));
        Assert.Equal("435", Text(pivot, "B4"));
        AssertValid(document);
    }

    [Fact]
    public void ColumnOnly_SumsPerColumnMemberAndGrandTotal()
    {
        using var document = CreateWorkbook(out var source, out var pivot);

        PivotTableBuilder.FromRange(source, Range).Column("Quarter").Value("Sales").PlaceAt(pivot, "A1");

        Assert.Equal("Q1", Text(pivot, "B1"));
        Assert.Equal("Q2", Text(pivot, "C1"));
        Assert.Equal("Grand Total", Text(pivot, "D1"));
        Assert.Equal("160", Text(pivot, "B2"));
        Assert.Equal("275", Text(pivot, "C2"));
        Assert.Equal("435", Text(pivot, "D2"));
        AssertValid(document);
    }

    [Fact]
    public void MultipleRowFields_NestsLabelsAndTotals()
    {
        using var document = CreateWorkbook(out var source, out var pivot);

        PivotTableBuilder.FromRange(source, Range)
            .Row("Region").Row("Channel").Column("Quarter").Value("Sales")
            .PlaceAt(pivot, "A1");

        Assert.Equal("East", Text(pivot, "A2"));
        Assert.Equal("Online", Text(pivot, "B2"));
        Assert.Equal("Retail", Text(pivot, "B3"));
        Assert.Equal("West", Text(pivot, "A4"));
        Assert.Equal("100", Text(pivot, "C2"));  // East/Online/Q1
        Assert.Equal("200", Text(pivot, "D2"));  // East/Online/Q2
        Assert.Equal("300", Text(pivot, "E2"));  // East/Online row total
        Assert.Equal("10", Text(pivot, "C3"));   // East/Retail/Q1
        Assert.Equal("Grand Total", Text(pivot, "A6"));
        Assert.Equal("160", Text(pivot, "C6"));
        Assert.Equal("435", Text(pivot, "E6"));
        AssertValid(document);
    }

    [Fact]
    public void MultipleValueFields_PlacesValueSelectorOnColumns()
    {
        using var document = CreateWorkbook(out var source, out var pivot);

        PivotTableBuilder.FromRange(source, Range)
            .Row("Region").Column("Quarter")
            .Value("Sales").Value("Units")
            .PlaceAt(pivot, "A1");

        Assert.Equal("Row Labels", Text(pivot, "A1"));
        Assert.Equal("Q1", Text(pivot, "B1"));
        Assert.Equal("Q2", Text(pivot, "D1"));
        Assert.Equal("Sum of Sales", Text(pivot, "B2"));
        Assert.Equal("Sum of Units", Text(pivot, "C2"));
        Assert.Equal("110", Text(pivot, "B3"));  // East/Q1 sales
        Assert.Equal("6", Text(pivot, "C3"));    // East/Q1 units
        Assert.Equal("200", Text(pivot, "D3"));  // East/Q2 sales
        Assert.Equal("310", Text(pivot, "F3"));  // East sales grand
        Assert.Equal("14", Text(pivot, "G3"));   // East units grand
        Assert.Equal("Grand Total", Text(pivot, "A5"));
        Assert.Equal("435", Text(pivot, "F5"));
        Assert.Equal("21", Text(pivot, "G5"));
        AssertValid(document);
    }

    [Fact]
    public void FilterSelection_RestrictsRowsAndRendersPageField()
    {
        using var document = CreateWorkbook(out var source, out var pivot);

        PivotTableBuilder.FromRange(source, Range)
            .Filter("Region", "East")
            .Row("Channel").Column("Quarter").Value("Sales")
            .PlaceAt(pivot, "A1");

        // Page-field row at the top.
        Assert.Equal("Region", Text(pivot, "A1"));
        Assert.Equal("East", Text(pivot, "B1"));

        // Body is shifted below the page area; only East rows are aggregated.
        Assert.Equal("Online", Text(pivot, "A4"));
        Assert.Equal("100", Text(pivot, "B4"));
        Assert.Equal("200", Text(pivot, "C4"));
        Assert.Equal("300", Text(pivot, "D4"));
        Assert.Equal("Retail", Text(pivot, "A5"));
        Assert.Equal("10", Text(pivot, "B5"));
        Assert.Equal("Grand Total", Text(pivot, "A6"));
        Assert.Equal("110", Text(pivot, "B6"));
        Assert.Equal("310", Text(pivot, "D6"));  // East-only grand, not 435
        AssertValid(document);
    }

    [Fact]
    public void ShowAsPercentOfColumn_RendersFractionsAndNormalizesTotals()
    {
        using var document = CreateWorkbook(out var source, out var pivot);

        PivotTableBuilder.FromRange(source, Range)
            .Row("Region").Column("Quarter")
            .Value("Sales", PivotAggregate.Sum, showAs: PivotShowAs.PercentOfColumn)
            .PlaceAt(pivot, "A1");

        // Q1 column total is 160; East/Q1 = 110/160 = 0.6875, West/Q1 = 50/160 = 0.3125.
        Assert.Equal("0.6875", Text(pivot, "B2"));
        Assert.Equal("0.3125", Text(pivot, "B3"));

        // Each data column sums to 100% in the grand-total row.
        Assert.Equal("1", Text(pivot, "B4"));
        Assert.Equal("1", Text(pivot, "C4"));
        Assert.Equal("1", Text(pivot, "D4"));

        var dataField = pivot.PivotTableParts.Single().PivotTableDefinition!
            .Descendants<DataField>().Single();
        Assert.Equal(ShowDataAsValues.PercentOfColumn, dataField.ShowDataAs!.Value);
        Assert.Equal(10u, dataField.NumberFormatId!.Value);
        AssertValid(document);
    }

    private static SpreadsheetDocument CreateWorkbook(out WorksheetPart source, out WorksheetPart pivot)
    {
        var stream = new MemoryStream();
        var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        source = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData();
        source.Worksheet = new Worksheet(data);
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(source), SheetId = 1, Name = "Data" });

        pivot = workbookPart.AddNewPart<WorksheetPart>();
        pivot.Worksheet = new Worksheet(new SheetData());
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(pivot), SheetId = 2, Name = "Pivot" });

        Header(data, "Region", "Channel", "Quarter", "Sales", "Units");
        DataRow(data, 2, "East", "Online", "Q1", 100, 5);
        DataRow(data, 3, "East", "Online", "Q2", 200, 8);
        DataRow(data, 4, "East", "Retail", "Q1", 10, 1);
        DataRow(data, 5, "West", "Online", "Q1", 50, 3);
        DataRow(data, 6, "West", "Retail", "Q2", 75, 4);
        return document;
    }

    private static void Header(SheetData sheetData, params string[] names)
    {
        var row = new Row { RowIndex = 1u };
        for (var i = 0; i < names.Length; i++)
        {
            row.Append(StringCell(PivotRefColumn(i + 1) + "1", names[i]));
        }

        sheetData.Append(row);
    }

    private static void DataRow(SheetData sheetData, uint rowIndex, string region, string channel, string quarter, double sales, double units)
    {
        var row = new Row { RowIndex = rowIndex };
        row.Append(StringCell("A" + rowIndex, region));
        row.Append(StringCell("B" + rowIndex, channel));
        row.Append(StringCell("C" + rowIndex, quarter));
        row.Append(NumberCell("D" + rowIndex, sales));
        row.Append(NumberCell("E" + rowIndex, units));
        sheetData.Append(row);
    }

    private static Cell StringCell(string reference, string value) => new()
    {
        CellReference = reference,
        DataType = CellValues.String,
        CellValue = new CellValue(value),
    };

    private static Cell NumberCell(string reference, double value) => new()
    {
        CellReference = reference,
        CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture)),
    };

    private static string PivotRefColumn(int column)
    {
        var result = string.Empty;
        while (column > 0)
        {
            var modulo = (column - 1) % 26;
            result = (char)('A' + modulo) + result;
            column = (column - modulo) / 26;
        }

        return result;
    }

    private static string Text(WorksheetPart part, string reference)
    {
        var cell = part.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference is not null && c.CellReference == reference);
        Assert.NotNull(cell);
        return cell!.InnerText;
    }

    private static void AssertValid(SpreadsheetDocument document)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(document).ToList();
        Assert.True(errors.Count == 0, "Validation errors:\n" + string.Join("\n", errors.Select(e => e.Id + " " + e.Description + " @ " + e.Path?.XPath)));
    }
}
