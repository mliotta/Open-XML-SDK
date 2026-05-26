// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Writes the precomputed pivot grid (nested labels, values, and grand totals) into the host
/// worksheet cells so the result is visible without recalculation.
/// </summary>
internal static class PivotSheetWriter
{
    /// <summary>Renders the model into the worksheet.</summary>
    /// <param name="targetWorksheetPart">The worksheet to render into.</param>
    /// <param name="model">The computed grid.</param>
    /// <param name="plan">The validated plan.</param>
    public static void Write(WorksheetPart targetWorksheetPart, PivotModel model, PivotPlan plan)
    {
        var g = PivotGeometry.Create(model, plan);
        var worksheet = targetWorksheetPart.Worksheet ?? throw new System.InvalidOperationException("The target worksheet has no content.");
        var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

        WritePageFields(sheetData, g, plan);
        WriteCorner(sheetData, g, model);
        WriteColumnHeaders(sheetData, g, model);
        WriteValueHeaders(sheetData, g, model);
        WriteRowLabels(sheetData, g, model);
        WriteBody(sheetData, g, model);
        WriteGrandRow(sheetData, g, model);
    }

    private static void WritePageFields(SheetData sheetData, PivotGeometry g, PivotPlan plan)
    {
        for (var i = 0; i < g.PageFieldCount; i++)
        {
            var filter = plan.Filters[i];
            SetString(sheetData, g.OriginColumn, g.OriginRow + i, filter.FieldName);
            SetString(sheetData, g.OriginColumn + 1, g.OriginRow + i, filter.SelectedValue ?? "(All)");
        }
    }

    private static void WriteCorner(SheetData sheetData, PivotGeometry g, PivotModel model)
    {
        var caption = g.ValueCount == 1 ? model.ValueDisplayNames[0] : "Row Labels";
        SetString(sheetData, g.OriginColumn, g.BodyOriginRow, caption);
    }

    private static void WriteColumnHeaders(SheetData sheetData, PivotGeometry g, PivotModel model)
    {
        for (var ct = 0; ct < g.ColumnTupleCount; ct++)
        {
            var firstDiff = ct == 0 ? 0 : FirstDifferingLevel(model.ColumnTuples[ct - 1], model.ColumnTuples[ct]);
            for (var f = firstDiff; f < g.ColumnFieldCount; f++)
            {
                var member = model.ColumnFieldMembers[f][model.ColumnTuples[ct][f]];
                SetString(sheetData, g.DataColumn(ct, 0), g.BodyOriginRow + f, member);
            }
        }

        if (g.ShowGrandColumn)
        {
            SetString(sheetData, g.FirstGrandColumn, g.BodyOriginRow, "Grand Total");
        }
    }

    private static void WriteValueHeaders(SheetData sheetData, PivotGeometry g, PivotModel model)
    {
        if (!g.HasValueHeaderRow)
        {
            return;
        }

        var row = g.ValueHeaderRow;
        for (var ct = 0; ct < g.ColumnTupleCount; ct++)
        {
            for (var v = 0; v < g.ValueCount; v++)
            {
                SetString(sheetData, g.DataColumn(ct, v), row, model.ValueDisplayNames[v]);
            }
        }

        if (g.ShowGrandColumn)
        {
            for (var v = 0; v < g.ValueCount; v++)
            {
                SetString(sheetData, g.FirstGrandColumn + v, row, model.ValueDisplayNames[v]);
            }
        }
    }

    private static void WriteRowLabels(SheetData sheetData, PivotGeometry g, PivotModel model)
    {
        for (var rt = 0; rt < g.RowTupleCount; rt++)
        {
            var firstDiff = rt == 0 ? 0 : FirstDifferingLevel(model.RowTuples[rt - 1], model.RowTuples[rt]);
            for (var f = firstDiff; f < g.RowFieldCount; f++)
            {
                var member = model.RowFieldMembers[f][model.RowTuples[rt][f]];
                SetString(sheetData, g.OriginColumn + f, g.FirstDataRow + rt, member);
            }
        }
    }

    private static void WriteBody(SheetData sheetData, PivotGeometry g, PivotModel model)
    {
        for (var rt = 0; rt < g.RowTupleCount; rt++)
        {
            var dataRow = g.FirstDataRow + rt;
            for (var ct = 0; ct < g.ColumnTupleCount; ct++)
            {
                for (var v = 0; v < g.ValueCount; v++)
                {
                    SetNumber(sheetData, g.DataColumn(ct, v), dataRow, model.Data[rt][ct][v]);
                }
            }

            if (g.ShowGrandColumn)
            {
                for (var v = 0; v < g.ValueCount; v++)
                {
                    SetNumber(sheetData, g.FirstGrandColumn + v, dataRow, model.RowTotals[rt][v]);
                }
            }
        }
    }

    private static void WriteGrandRow(SheetData sheetData, PivotGeometry g, PivotModel model)
    {
        if (!g.ShowGrandRow)
        {
            return;
        }

        var row = g.GrandRow;
        SetString(sheetData, g.OriginColumn, row, "Grand Total");
        for (var ct = 0; ct < g.ColumnTupleCount; ct++)
        {
            for (var v = 0; v < g.ValueCount; v++)
            {
                SetNumber(sheetData, g.DataColumn(ct, v), row, model.ColumnTotals[ct][v]);
            }
        }

        if (g.ShowGrandColumn)
        {
            for (var v = 0; v < g.ValueCount; v++)
            {
                SetNumber(sheetData, g.FirstGrandColumn + v, row, model.GrandTotals[v]);
            }
        }
    }

    private static int FirstDifferingLevel(int[] previous, int[] current)
    {
        for (var i = 0; i < current.Length; i++)
        {
            if (i >= previous.Length || previous[i] != current[i])
            {
                return i;
            }
        }

        return current.Length;
    }

    private static void SetString(SheetData sheetData, int column, int row, string text)
    {
        var cell = GetCell(sheetData, column, row);
        cell.RemoveAllChildren();
        cell.DataType = CellValues.InlineString;
        cell.Append(new InlineString(new Text(text)));
    }

    private static void SetNumber(SheetData sheetData, int column, int row, double? value)
    {
        if (value is null)
        {
            return;
        }

        var cell = GetCell(sheetData, column, row);
        cell.RemoveAllChildren();
        cell.DataType = null;
        cell.Append(new CellValue(value.Value.ToString(CultureInfo.InvariantCulture)));
    }

    private static Cell GetCell(SheetData sheetData, int column, int row)
    {
        var rowElement = GetRow(sheetData, (uint)row);
        var reference = PivotRef.Cell(column, row);

        foreach (var existing in rowElement.Elements<Cell>())
        {
            if (existing.CellReference is not null && existing.CellReference == reference)
            {
                return existing;
            }
        }

        var cell = new Cell { CellReference = reference };
        var successor = rowElement.Elements<Cell>()
            .FirstOrDefault(c => c.CellReference?.Value is not null && PivotRef.ParseCell(c.CellReference!.Value!).Column > column);
        if (successor is not null)
        {
            rowElement.InsertBefore(cell, successor);
        }
        else
        {
            rowElement.Append(cell);
        }

        return cell;
    }

    private static Row GetRow(SheetData sheetData, uint row)
    {
        foreach (var existing in sheetData.Elements<Row>())
        {
            if (existing.RowIndex is not null && existing.RowIndex == row)
            {
                return existing;
            }
        }

        var rowElement = new Row { RowIndex = row };
        var successor = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value is not null && r.RowIndex!.Value > row);
        if (successor is not null)
        {
            sheetData.InsertBefore(rowElement, successor);
        }
        else
        {
            sheetData.Append(rowElement);
        }

        return rowElement;
    }
}
