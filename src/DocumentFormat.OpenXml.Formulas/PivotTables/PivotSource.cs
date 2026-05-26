// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Features.FormulaEvaluation;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>Reads a worksheet range into a <see cref="PivotSourceData"/>.</summary>
internal static class PivotSource
{
    /// <summary>Reads the source range, treating its first row as headers.</summary>
    /// <param name="sourceWorksheetPart">The worksheet part holding the data.</param>
    /// <param name="workbookPart">The owning workbook part.</param>
    /// <param name="reference">The source range reference.</param>
    /// <param name="context">A cell context bound to the source worksheet.</param>
    /// <returns>The parsed source data.</returns>
    public static PivotSourceData Read(WorksheetPart sourceWorksheetPart, WorkbookPart workbookPart, string reference, CellContext context)
    {
        var bounds = PivotRef.ParseRange(reference);

        var width = bounds.LastColumn - bounds.FirstColumn + 1;
        var headers = new string[width];
        for (var c = 0; c < width; c++)
        {
            var value = context.GetCell(PivotRef.Cell(bounds.FirstColumn + c, bounds.FirstRow));
            headers[c] = ToText(value);
        }

        var rows = new List<FormulaResult[]>();
        for (var r = bounds.FirstRow + 1; r <= bounds.LastRow; r++)
        {
            var row = new FormulaResult[width];
            for (var c = 0; c < width; c++)
            {
                row[c] = context.GetCell(PivotRef.Cell(bounds.FirstColumn + c, r));
            }

            rows.Add(row);
        }

        return new PivotSourceData(ResolveSheetName(workbookPart, sourceWorksheetPart), reference, headers, rows);
    }

    /// <summary>Renders a result as the text Excel would display (invariant culture).</summary>
    /// <param name="value">The result.</param>
    /// <returns>The display text.</returns>
    public static string ToText(FormulaResult value) => value.Type switch
    {
        FormulaResultType.Number => value.NumericValue.ToString(CultureInfo.InvariantCulture),
        FormulaResultType.Boolean => value.BoolValue ? "TRUE" : "FALSE",
        FormulaResultType.Error => value.ErrorValue ?? "#VALUE!",
        FormulaResultType.Empty => string.Empty,
        _ => value.StringValue,
    };

    private static string ResolveSheetName(WorkbookPart workbookPart, WorksheetPart worksheetPart)
    {
        var id = workbookPart.GetIdOfPart(worksheetPart);
        var sheet = workbookPart.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault(s => s.Id is not null && s.Id == id);
        return sheet?.Name?.Value ?? "Sheet1";
    }
}
