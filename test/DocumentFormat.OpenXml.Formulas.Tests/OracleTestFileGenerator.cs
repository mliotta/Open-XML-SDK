// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Tests;

/// <summary>
/// Generates Excel files with comprehensive formula test cases for oracle validation.
/// </summary>
public static class OracleTestFileGenerator
{
    /// <summary>
    /// Generates an Excel file with comprehensive formula test cases.
    /// User should open this file in Excel, which will calculate and cache all formula results.
    /// </summary>
    public static void GenerateOracleTestFile(string filePath)
    {
        using var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        // Create multiple sheets for different function categories
        CreateMathFunctionsSheet(workbookPart, sheets, 1);
        CreateLogicalFunctionsSheet(workbookPart, sheets, 2);
        CreateTextFunctionsSheet(workbookPart, sheets, 3);
        CreateLookupFunctionsSheet(workbookPart, sheets, 4);
        CreateDateTimeFunctionsSheet(workbookPart, sheets, 5);
        CreateStatisticalFunctionsSheet(workbookPart, sheets, 6);
        CreateInformationFunctionsSheet(workbookPart, sheets, 7);

        workbookPart.Workbook.Save();
    }

    private static void CreateMathFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Math",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // SUM tests
        AddTestCase(sheetData, row++, "SUM", "=SUM(1,2,3)", "Basic sum");
        AddTestCase(sheetData, row++, "SUM", "=SUM(F3:F5)", "Range sum", ("F3", "10"), ("G3", "20"), ("H3", "30"));
        AddTestCase(sheetData, row++, "SUM", "=SUM(1,2,F4)", "Mixed sum", ("F4", "5"));

        // AVERAGE tests
        AddTestCase(sheetData, row++, "AVERAGE", "=AVERAGE(10,20,30)", "Basic average");
        AddTestCase(sheetData, row++, "AVERAGE", "=AVERAGE(F6:H6)", "Range average", ("F6", "5"), ("G6", "15"), ("H6", "25"));

        // COUNT tests
        AddTestCase(sheetData, row++, "COUNT", "=COUNT(1,2,3,4,5)", "Basic count");
        AddTestCase(sheetData, row++, "COUNT", "=COUNT(F8:J8)", "Range count", ("F8", "1"), ("G8", "2"), ("H8", "3"), ("I8", "4"), ("J8", "5"));

        // COUNTA tests
        AddTestCase(sheetData, row++, "COUNTA", "=COUNTA(1,\"text\",TRUE)", "Count with mixed types");

        // MAX/MIN tests
        AddTestCase(sheetData, row++, "MAX", "=MAX(5,10,3,8)", "Basic max");
        AddTestCase(sheetData, row++, "MIN", "=MIN(5,10,3,8)", "Basic min");
        AddTestCase(sheetData, row++, "MAX", "=MAX(F12:I12)", "Range max", ("F12", "15"), ("G12", "25"), ("H12", "10"), ("I12", "20"));

        // ROUND tests
        AddTestCase(sheetData, row++, "ROUND", "=ROUND(3.14159, 2)", "Round to 2 decimals");
        AddTestCase(sheetData, row++, "ROUND", "=ROUND(2.5, 0)", "Round half up");
        AddTestCase(sheetData, row++, "ROUNDUP", "=ROUNDUP(3.14, 1)", "Round up");
        AddTestCase(sheetData, row++, "ROUNDDOWN", "=ROUNDDOWN(3.99, 1)", "Round down");

        // ABS tests
        AddTestCase(sheetData, row++, "ABS", "=ABS(-10)", "Absolute value negative");
        AddTestCase(sheetData, row++, "ABS", "=ABS(10)", "Absolute value positive");

        // PRODUCT tests
        AddTestCase(sheetData, row++, "PRODUCT", "=PRODUCT(2,3,4)", "Basic product");
        AddTestCase(sheetData, row++, "PRODUCT", "=PRODUCT(F20:H20)", "Range product", ("F20", "2"), ("G20", "5"), ("H20", "10"));

        // POWER tests
        AddTestCase(sheetData, row++, "POWER", "=POWER(2,3)", "2 to the power of 3");
        AddTestCase(sheetData, row++, "POWER", "=POWER(5,2)", "5 squared");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateLogicalFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Logical",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // IF tests
        AddTestCase(sheetData, row++, "IF", "=IF(TRUE, \"yes\", \"no\")", "IF true");
        AddTestCase(sheetData, row++, "IF", "=IF(FALSE, \"yes\", \"no\")", "IF false");
        AddTestCase(sheetData, row++, "IF", "=IF(10>5, \"greater\", \"less\")", "IF comparison");
        AddTestCase(sheetData, row++, "IF", "=IF(F5>10, \"high\", \"low\")", "IF cell reference", ("F5", "15"));

        // Nested IF
        AddTestCase(sheetData, row++, "IF", "=IF(F6>20, \"high\", IF(F6>10, \"medium\", \"low\"))", "Nested IF", ("F6", "15"));

        // AND tests
        AddTestCase(sheetData, row++, "AND", "=AND(TRUE, TRUE)", "AND both true");
        AddTestCase(sheetData, row++, "AND", "=AND(TRUE, FALSE)", "AND one false");
        AddTestCase(sheetData, row++, "AND", "=AND(10>5, 20>15)", "AND comparisons");

        // OR tests
        AddTestCase(sheetData, row++, "OR", "=OR(TRUE, FALSE)", "OR one true");
        AddTestCase(sheetData, row++, "OR", "=OR(FALSE, FALSE)", "OR both false");
        AddTestCase(sheetData, row++, "OR", "=OR(10>5, 3>15)", "OR comparisons");

        // NOT tests
        AddTestCase(sheetData, row++, "NOT", "=NOT(TRUE)", "NOT true");
        AddTestCase(sheetData, row++, "NOT", "=NOT(FALSE)", "NOT false");

        // Complex logical expressions
        AddTestCase(sheetData, row++, "AND+OR", "=AND(OR(F15>10, F15<5), F15<>0)", "Complex logic", ("F15", "12"));

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateTextFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Text",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Setup test data in columns F, G, H
        AddTextCell(sheetData, "F1", "Hello");
        AddTextCell(sheetData, "G1", "World");
        AddTextCell(sheetData, "H1", "  Excel  ");

        // CONCATENATE tests
        AddTestCase(sheetData, row++, "CONCATENATE", "=CONCATENATE(\"Hello\", \" \", \"World\")", "Basic concatenate");
        AddTestCase(sheetData, row++, "CONCATENATE", "=CONCATENATE(F1, \" \", G1)", "Concatenate cells");

        // LEFT/RIGHT/MID tests
        AddTestCase(sheetData, row++, "LEFT", "=LEFT(\"Hello\", 3)", "Left 3 chars");
        AddTestCase(sheetData, row++, "RIGHT", "=RIGHT(\"World\", 3)", "Right 3 chars");
        AddTestCase(sheetData, row++, "MID", "=MID(\"Hello\", 2, 3)", "Mid substring");

        // LEN tests
        AddTestCase(sheetData, row++, "LEN", "=LEN(\"Hello\")", "Length of string");
        AddTestCase(sheetData, row++, "LEN", "=LEN(F1)", "Length of cell");

        // TRIM tests
        AddTestCase(sheetData, row++, "TRIM", "=TRIM(\"  Hello  \")", "Trim spaces");
        AddTestCase(sheetData, row++, "TRIM", "=TRIM(H1)", "Trim cell");

        // UPPER/LOWER/PROPER tests
        AddTestCase(sheetData, row++, "UPPER", "=UPPER(\"hello\")", "Convert to uppercase");
        AddTestCase(sheetData, row++, "LOWER", "=LOWER(\"WORLD\")", "Convert to lowercase");
        AddTestCase(sheetData, row++, "PROPER", "=PROPER(\"hello world\")", "Convert to proper case");

        // TEXT tests
        AddTestCase(sheetData, row++, "TEXT", "=TEXT(1234.5, \"0.00\")", "Format number");
        AddTestCase(sheetData, row++, "TEXT", "=TEXT(0.75, \"0%\")", "Format as percent");

        // VALUE tests
        AddTestCase(sheetData, row++, "VALUE", "=VALUE(\"123\")", "Convert string to number");
        AddTestCase(sheetData, row++, "VALUE", "=VALUE(\"45.67\")", "Convert decimal string");

        // FIND/SEARCH tests
        AddTestCase(sheetData, row++, "FIND", "=FIND(\"l\", \"Hello\")", "Find character");
        AddTestCase(sheetData, row++, "SEARCH", "=SEARCH(\"o\", \"Hello\")", "Search character");

        // SUBSTITUTE tests
        AddTestCase(sheetData, row++, "SUBSTITUTE", "=SUBSTITUTE(\"Hello\", \"l\", \"L\")", "Substitute all");
        AddTestCase(sheetData, row++, "SUBSTITUTE", "=SUBSTITUTE(\"banana\", \"a\", \"o\", 2)", "Substitute 2nd occurrence");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateLookupFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Lookup",
        });

        // Create lookup table (A1:C5) - note: no standard header for this sheet
        AddTextCell(sheetData, "A1", "ID");
        AddTextCell(sheetData, "B1", "Name");
        AddTextCell(sheetData, "C1", "Score");

        AddCell(sheetData, "A2", "1");
        AddTextCell(sheetData, "B2", "Alice");
        AddCell(sheetData, "C2", "95");

        AddCell(sheetData, "A3", "2");
        AddTextCell(sheetData, "B3", "Bob");
        AddCell(sheetData, "C3", "87");

        AddCell(sheetData, "A4", "3");
        AddTextCell(sheetData, "B4", "Carol");
        AddCell(sheetData, "C4", "92");

        AddCell(sheetData, "A5", "4");
        AddTextCell(sheetData, "B5", "Dave");
        AddCell(sheetData, "C5", "88");

        // Test cases start at row 7 (after lookup table data)
        var row = 7;

        // VLOOKUP tests
        AddTestCase(sheetData, row++, "VLOOKUP", "=VLOOKUP(2, A2:C5, 2, FALSE)", "VLOOKUP name by ID");
        AddTestCase(sheetData, row++, "VLOOKUP", "=VLOOKUP(3, A2:C5, 3, FALSE)", "VLOOKUP score by ID");
        AddTestCase(sheetData, row++, "VLOOKUP", "=VLOOKUP(1, A2:C5, 2, FALSE)", "VLOOKUP first row");

        // HLOOKUP tests (create horizontal table)
        AddTextCell(sheetData, "F1", "Alice");
        AddTextCell(sheetData, "G1", "Bob");
        AddTextCell(sheetData, "H1", "Carol");
        AddCell(sheetData, "F2", "95");
        AddCell(sheetData, "G2", "87");
        AddCell(sheetData, "H2", "92");

        AddTestCase(sheetData, row++, "HLOOKUP", "=HLOOKUP(\"Bob\", F1:H2, 2, FALSE)", "HLOOKUP score by name");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateDateTimeFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "DateTime",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Store test data in column F (avoid conflict with test case columns A-D)
        AddCell(sheetData, "F1", "45096"); // Excel serial date for 2023-06-15
        AddCell(sheetData, "G1", "0.5"); // Noon (12:00:00)

        // DATE tests
        AddTestCase(sheetData, row++, "DATE", "=DATE(2023, 6, 15)", "Create date");
        AddTestCase(sheetData, row++, "DATE", "=DATE(2024, 1, 1)", "New year date");

        // YEAR/MONTH/DAY tests
        AddTestCase(sheetData, row++, "YEAR", "=YEAR(F1)", "Extract year");
        AddTestCase(sheetData, row++, "MONTH", "=MONTH(F1)", "Extract month");
        AddTestCase(sheetData, row++, "DAY", "=DAY(F1)", "Extract day");

        // WEEKDAY test
        AddTestCase(sheetData, row++, "WEEKDAY", "=WEEKDAY(F1)", "Get weekday");

        // TODAY/NOW tests (will vary by when Excel calculates)
        AddTestCase(sheetData, row++, "TODAY", "=YEAR(TODAY())", "Today's year");
        AddTestCase(sheetData, row++, "NOW", "=YEAR(NOW())", "Now's year");

        // HOUR/MINUTE/SECOND tests (using time value)
        AddTestCase(sheetData, row++, "HOUR", "=HOUR(G1)", "Extract hour from time");
        AddTestCase(sheetData, row++, "MINUTE", "=MINUTE(G1)", "Extract minute from time");
        AddTestCase(sheetData, row++, "SECOND", "=SECOND(G1)", "Extract second from time");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateStatisticalFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Statistical",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Create test data in column F (avoid conflict with test case columns A-D)
        for (int i = 0; i < 10; i++)
        {
            AddCell(sheetData, $"F{i + 1}", ((i + 1) * 5).ToString());
        }

        // MEDIAN tests
        AddTestCase(sheetData, row++, "MEDIAN", "=MEDIAN(1,2,3,4,5)", "Median odd count");
        AddTestCase(sheetData, row++, "MEDIAN", "=MEDIAN(1,2,3,4)", "Median even count");
        AddTestCase(sheetData, row++, "MEDIAN", "=MEDIAN(F1:F10)", "Median range");

        // MODE tests
        AddTestCase(sheetData, row++, "MODE", "=MODE(1,2,2,3,4)", "Mode basic");

        // STDEV tests
        AddTestCase(sheetData, row++, "STDEV", "=STDEV(1,2,3,4,5)", "Standard deviation");
        AddTestCase(sheetData, row++, "STDEV", "=STDEV(F1:F5)", "Stdev range");

        // VAR tests
        AddTestCase(sheetData, row++, "VAR", "=VAR(1,2,3,4,5)", "Variance");
        AddTestCase(sheetData, row++, "VAR", "=VAR(F1:F5)", "Var range");

        // RANK tests
        AddTestCase(sheetData, row++, "RANK", "=RANK(25, F1:F10)", "Rank in range");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void CreateInformationFunctionsSheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Information",
        });

        AddHeader(sheetData, 1);
        var row = 2;

        // Setup test data in column F (avoid conflict with test case columns A-D)
        AddCell(sheetData, "F1", "123");
        AddTextCell(sheetData, "F2", "Hello");

        // ISNUMBER tests
        AddTestCase(sheetData, row++, "ISNUMBER", "=ISNUMBER(123)", "Is number literal");
        AddTestCase(sheetData, row++, "ISNUMBER", "=ISNUMBER(\"text\")", "Is number text");
        AddTestCase(sheetData, row++, "ISNUMBER", "=ISNUMBER(F1)", "Is number cell");

        // ISTEXT tests
        AddTestCase(sheetData, row++, "ISTEXT", "=ISTEXT(\"Hello\")", "Is text literal");
        AddTestCase(sheetData, row++, "ISTEXT", "=ISTEXT(123)", "Is text number");
        AddTestCase(sheetData, row++, "ISTEXT", "=ISTEXT(F2)", "Is text cell");

        SortCellsInRows(sheetData);
        worksheetPart.Worksheet.Save();
    }

    private static void AddHeader(SheetData sheetData, uint rowIndex)
    {
        var row = new Row { RowIndex = rowIndex };
        sheetData.AppendChild(row);

        row.AppendChild(new Cell
        {
            CellReference = $"A{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Function"),
        });

        row.AppendChild(new Cell
        {
            CellReference = $"B{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Formula"),
        });

        row.AppendChild(new Cell
        {
            CellReference = $"C{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Expected (Excel)"),
        });

        row.AppendChild(new Cell
        {
            CellReference = $"D{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Description"),
        });
    }

    private static void AddTestCase(SheetData sheetData, int rowIndex, string functionName, string formula, string description, params (string CellRef, string Value)[] setupData)
    {
        // Add setup data first
        foreach (var (cellRef, value) in setupData)
        {
            AddCell(sheetData, cellRef, value);
        }

        // Get or create the row for this test case
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            sheetData.AppendChild(row);
        }

        // Column A: Function name
        row.AppendChild(new Cell
        {
            CellReference = $"A{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(functionName),
        });

        // Column B: Formula (as text for reference)
        row.AppendChild(new Cell
        {
            CellReference = $"B{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(formula),
        });

        // Column C: Formula (actual formula for Excel to calculate)
        row.AppendChild(new Cell
        {
            CellReference = $"C{rowIndex}",
            CellFormula = new CellFormula(formula),
            // No CellValue - Excel will calculate
        });

        // Column D: Description
        row.AppendChild(new Cell
        {
            CellReference = $"D{rowIndex}",
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(description),
        });
    }

    private static void AddCell(SheetData sheetData, string cellRef, string value)
    {
        var rowIndex = GetRowIndex(cellRef);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.AppendChild(row);
        }

        row.AppendChild(new Cell
        {
            CellReference = cellRef,
            DataType = CellValues.Number,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value),
        });
    }

    private static void AddTextCell(SheetData sheetData, string cellRef, string value)
    {
        var rowIndex = GetRowIndex(cellRef);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            sheetData.AppendChild(row);
        }

        row.AppendChild(new Cell
        {
            CellReference = cellRef,
            DataType = CellValues.String,
            CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(value),
        });
    }

    private static uint GetRowIndex(string reference)
    {
        return uint.Parse(new string(reference.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Sorts cells within each row by cell reference (required by Excel).
    /// Excel requires cells to be in alphabetical order within a row (A, B, C... not C, A, B).
    /// </summary>
    private static void SortCellsInRows(SheetData sheetData)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            var cells = row.Elements<Cell>().ToList();

            // Sort by cell reference
            var sortedCells = cells.OrderBy(c => c.CellReference?.Value).ToList();

            // Remove all cells
            row.RemoveAllChildren<Cell>();

            // Add back in sorted order
            foreach (var cell in sortedCells)
            {
                row.AppendChild(cell);
            }
        }
    }
}
