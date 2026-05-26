// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

using Xunit;

namespace DocumentFormat.OpenXml.Features.PivotTables.Tests;

/// <summary>
/// End-to-end tests for <see cref="PivotTableBuilder"/>: precomputed values plus a schema-valid
/// package.
/// </summary>
public class PivotTableBuilderTests
{
    [Fact]
    public void PlaceAt_RowColumnSumPivot_WritesPrecomputedValuesAndValidates()
    {
        using var stream = new MemoryStream();
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        var sourcePart = workbookPart.AddNewPart<WorksheetPart>();
        var sourceData = new SheetData();
        sourcePart.Worksheet = new Worksheet(sourceData);
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(sourcePart), SheetId = 1, Name = "Data" });

        var pivotPart = workbookPart.AddNewPart<WorksheetPart>();
        pivotPart.Worksheet = new Worksheet(new SheetData());
        sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(pivotPart), SheetId = 2, Name = "Pivot" });

        // Region | Quarter | Sales
        AddText(sourceData, "A1", "Region");
        AddText(sourceData, "B1", "Quarter");
        AddText(sourceData, "C1", "Sales");
        AddRow(sourceData, 2, "East", "Q1", 100);
        AddRow(sourceData, 3, "East", "Q2", 200);
        AddRow(sourceData, 4, "West", "Q1", 50);
        AddRow(sourceData, 5, "West", "Q2", 75);
        AddRow(sourceData, 6, "East", "Q1", 10);

        PivotTableBuilder
            .FromRange(sourcePart, "A1:C6")
            .Named("SalesByRegion")
            .Row("Region")
            .Column("Quarter")
            .Value("Sales", PivotAggregate.Sum)
            .PlaceAt(pivotPart, "A1");

        // Header row: A1 caption, B1=Q1, C1=Q2, D1=Grand Total.
        Assert.Equal("Sum of Sales", Text(pivotPart, "A1"));
        Assert.Equal("Q1", Text(pivotPart, "B1"));
        Assert.Equal("Q2", Text(pivotPart, "C1"));
        Assert.Equal("Grand Total", Text(pivotPart, "D1"));

        // East row: 110, 200, total 310.
        Assert.Equal("East", Text(pivotPart, "A2"));
        Assert.Equal("110", Text(pivotPart, "B2"));
        Assert.Equal("200", Text(pivotPart, "C2"));
        Assert.Equal("310", Text(pivotPart, "D2"));

        // West row: 50, 75, total 125.
        Assert.Equal("West", Text(pivotPart, "A3"));
        Assert.Equal("50", Text(pivotPart, "B3"));
        Assert.Equal("125", Text(pivotPart, "D3"));

        // Grand total row.
        Assert.Equal("Grand Total", Text(pivotPart, "A4"));
        Assert.Equal("160", Text(pivotPart, "B4"));
        Assert.Equal("275", Text(pivotPart, "C4"));
        Assert.Equal("435", Text(pivotPart, "D4"));

        // The cache definition and pivot parts exist and are linked.
        Assert.Single(workbookPart.PivotTableCacheDefinitionParts);
        Assert.Single(pivotPart.PivotTableParts);

        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(document).ToList();
        Assert.True(errors.Count == 0, "Validation errors:\n" + string.Join("\n", errors.Select(e => e.Id + " " + e.Description + " @ " + e.Path?.XPath)));
    }

    private static void AddRow(SheetData sheetData, uint rowIndex, string region, string quarter, double sales)
    {
        var row = new Row { RowIndex = rowIndex };
        row.Append(StringCell("A" + rowIndex, region));
        row.Append(StringCell("B" + rowIndex, quarter));
        row.Append(new Cell { CellReference = "C" + rowIndex, CellValue = new CellValue(sales.ToString(System.Globalization.CultureInfo.InvariantCulture)) });
        sheetData.Append(row);
    }

    private static void AddText(SheetData sheetData, string reference, string value)
    {
        var rowIndex = uint.Parse(new string(reference.Where(char.IsDigit).ToArray()), System.Globalization.CultureInfo.InvariantCulture);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex is not null && r.RowIndex == rowIndex)
            ?? (Row)sheetData.AppendChild(new Row { RowIndex = rowIndex });
        row.Append(StringCell(reference, value));
    }

    private static Cell StringCell(string reference, string value) => new()
    {
        CellReference = reference,
        DataType = CellValues.String,
        CellValue = new CellValue(value),
    };

    private static string Text(WorksheetPart part, string reference)
    {
        var cell = part.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference is not null && c.CellReference == reference);
        Assert.NotNull(cell);
        return cell!.InnerText;
    }
}
