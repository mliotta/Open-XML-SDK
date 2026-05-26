// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.Text;

namespace DocumentFormat.OpenXml.Features.PivotTables;

/// <summary>
/// Minimal A1-style reference arithmetic used while assembling pivot tables.
/// Columns and rows are one-based.
/// </summary>
internal static class PivotRef
{
    /// <summary>Parses a single cell reference such as <c>"B3"</c>.</summary>
    /// <param name="reference">The cell reference (any <c>$</c> anchors are ignored).</param>
    /// <returns>The one-based position.</returns>
    public static CellPosition ParseCell(string reference)
    {
        var clean = reference.Replace("$", string.Empty).Trim();
        var split = 0;
        while (split < clean.Length && char.IsLetter(clean[split]))
        {
            split++;
        }

        if (split == 0 || split == clean.Length)
        {
            throw new FormatException($"Invalid cell reference: '{reference}'.");
        }

        var column = 0;
        for (var i = 0; i < split; i++)
        {
            column = (column * 26) + (char.ToUpperInvariant(clean[i]) - 'A' + 1);
        }

        var row = int.Parse(clean.Substring(split), CultureInfo.InvariantCulture);
        return new CellPosition(column, row);
    }

    /// <summary>Parses a range such as <c>"A1:D100"</c> (or a single cell) into bounds.</summary>
    /// <param name="reference">The range reference.</param>
    /// <returns>The inclusive bounds.</returns>
    public static RangeBounds ParseRange(string reference)
    {
        var parts = reference.Split(':');
        var start = ParseCell(parts[0]);
        if (parts.Length == 1)
        {
            return new RangeBounds(start.Column, start.Row, start.Column, start.Row);
        }

        var end = ParseCell(parts[1]);
        return new RangeBounds(
            System.Math.Min(start.Column, end.Column),
            System.Math.Min(start.Row, end.Row),
            System.Math.Max(start.Column, end.Column),
            System.Math.Max(start.Row, end.Row));
    }

    /// <summary>Converts a one-based column index to its letter form (1 =&gt; "A").</summary>
    /// <param name="column">The one-based column index.</param>
    /// <returns>The column letters.</returns>
    public static string ColumnName(int column)
    {
        var builder = new StringBuilder();
        while (column > 0)
        {
            var modulo = (column - 1) % 26;
            builder.Insert(0, (char)('A' + modulo));
            column = (column - modulo) / 26;
        }

        return builder.ToString();
    }

    /// <summary>Builds a cell reference from one-based column and row.</summary>
    /// <param name="column">The one-based column index.</param>
    /// <param name="row">The one-based row index.</param>
    /// <returns>The cell reference, e.g. <c>"B3"</c>.</returns>
    public static string Cell(int column, int row) => ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

    /// <summary>Builds a range reference from one-based bounds.</summary>
    /// <param name="firstColumn">First column (one-based).</param>
    /// <param name="firstRow">First row (one-based).</param>
    /// <param name="lastColumn">Last column (one-based).</param>
    /// <param name="lastRow">Last row (one-based).</param>
    /// <returns>The range reference, e.g. <c>"A1:D100"</c>.</returns>
    public static string Range(int firstColumn, int firstRow, int lastColumn, int lastRow)
        => Cell(firstColumn, firstRow) + ":" + Cell(lastColumn, lastRow);
}
