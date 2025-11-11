// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

/// <summary>
/// Provides context for evaluating formulas, including cell value resolution.
/// </summary>
public class CellContext
{
    private readonly Worksheet _worksheet;
    private readonly SharedStringTablePart? _sharedStringTablePart;

    // TODO: Phase 0 limitation - cache never invalidates.
    // Phase 1 must add invalidation when cell values change.
    private readonly Dictionary<string, CellValue> _cache = new();

    /// <summary>
    /// Initializes a new instance of the <see cref="CellContext"/> class.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the cells.</param>
    /// <param name="sharedStringTablePart">The shared string table part for resolving shared strings.</param>
    public CellContext(Worksheet worksheet, SharedStringTablePart? sharedStringTablePart = null)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _sharedStringTablePart = sharedStringTablePart;
    }

    /// <summary>
    /// Gets the value of a cell by reference.
    /// </summary>
    /// <param name="reference">The cell reference (e.g., "A1").</param>
    /// <returns>The cell value.</returns>
    public CellValue GetCell(string reference)
    {
        if (_cache.TryGetValue(reference, out var cached))
        {
            return cached;
        }

        var cell = FindCell(_worksheet, reference);
        return _cache[reference] = ReadCellValue(cell);
    }

    /// <summary>
    /// Gets values for a range of cells.
    /// </summary>
    /// <param name="start">Start cell reference.</param>
    /// <param name="end">End cell reference.</param>
    /// <returns>Enumerable of cell values.</returns>
    public IEnumerable<CellValue> GetRange(string start, string end)
    {
        int startCol, startRow;
        ParseCellReference(start, out startCol, out startRow);

        int endCol, endRow;
        ParseCellReference(end, out endCol, out endRow);

        for (var row = startRow; row <= endRow; row++)
        {
            for (var col = startCol; col <= endCol; col++)
            {
                var cellRef = GetColumnLetter(col) + row.ToString(CultureInfo.InvariantCulture);
                yield return GetCell(cellRef);
            }
        }
    }

    private static Cell? FindCell(Worksheet worksheet, string reference)
    {
        var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData == null)
        {
            return null;
        }

        return sheetData.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
    }

    private CellValue ReadCellValue(Cell? cell)
    {
        if (cell == null)
        {
            return CellValue.Empty;
        }

        var cellValue = cell.CellValue?.Text;
        if (string.IsNullOrEmpty(cellValue))
        {
            return CellValue.Empty;
        }

        // Check data type
        var dataType = cell.DataType?.Value;

        if (dataType == CellValues.Boolean)
        {
            return CellValue.FromBool(cellValue == "1" || string.Equals(cellValue, "true", StringComparison.OrdinalIgnoreCase));
        }

        if (dataType == CellValues.Error)
        {
            return CellValue.Error(cellValue);
        }

        if (dataType == CellValues.SharedString)
        {
            // Resolve shared string index
            if (int.TryParse(cellValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index))
            {
                var sharedString = GetSharedString(index);
                if (sharedString != null)
                {
                    return CellValue.FromString(sharedString);
                }
            }

            // If we can't resolve, return the index as a string (fallback)
            return CellValue.FromString(cellValue);
        }

        if (dataType == CellValues.String || dataType == CellValues.InlineString)
        {
            return CellValue.FromString(cellValue);
        }

        // Try to parse as number
        if (double.TryParse(cellValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return CellValue.FromNumber(number);
        }

        return CellValue.FromString(cellValue);
    }

    private string? GetSharedString(int index)
    {
        if (_sharedStringTablePart == null)
        {
            return null;
        }

        var sharedStringTable = _sharedStringTablePart.SharedStringTable;
        if (sharedStringTable == null)
        {
            return null;
        }

        var items = sharedStringTable.Elements<SharedStringItem>().ToList();
        if (index >= 0 && index < items.Count)
        {
            // Get the text from the shared string item
            var item = items[index];
            return item.InnerText;
        }

        return null;
    }

    private static void ParseCellReference(string reference, out int column, out int row)
    {
        // Remove $ signs for absolute references
        reference = reference.Replace("$", string.Empty);

        var match = Regex.Match(reference, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            throw new ArgumentException($"Invalid cell reference: {reference}", nameof(reference));
        }

        var columnLetters = match.Groups[1].Value;
        row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);

        column = 0;
        for (var i = 0; i < columnLetters.Length; i++)
        {
            column = (column * 26) + (char.ToUpperInvariant(columnLetters[i]) - 'A' + 1);
        }
    }

    private static string GetColumnLetter(int column)
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
}
