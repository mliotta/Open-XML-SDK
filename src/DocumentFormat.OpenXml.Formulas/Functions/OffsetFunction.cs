// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the OFFSET function.
/// OFFSET(reference, rows, cols, [height], [width]) - Returns reference offset from base reference.
/// For Phase 0, returns the cell value at the offset position.
/// </summary>
public sealed class OffsetFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly OffsetFunction Instance = new();

    private OffsetFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "OFFSET";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract reference (required)
        var referenceArg = args[0];
        if (referenceArg.IsError)
        {
            return referenceArg;
        }

        // Get reference as string
        string reference;
        if (referenceArg.Type == CellValueType.Text)
        {
            reference = referenceArg.StringValue;
        }
        else
        {
            // If no explicit reference text, use current cell reference from context
            if (context?.CurrentCellReference == null)
            {
                return CellValue.Error("#VALUE!");
            }

            reference = context.CurrentCellReference;
        }

        // Parse the reference to get base row and column
        if (!TryParseCellReference(reference, out var baseCol, out var baseRow))
        {
            return CellValue.Error("#REF!");
        }

        // Extract rows offset (required)
        var rowsArg = args[1];
        if (rowsArg.IsError)
        {
            return rowsArg;
        }

        if (rowsArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rowsOffset = (int)rowsArg.NumericValue;

        // Extract cols offset (required)
        var colsArg = args[2];
        if (colsArg.IsError)
        {
            return colsArg;
        }

        if (colsArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var colsOffset = (int)colsArg.NumericValue;

        // Extract height (optional, default = 1)
        var height = 1;
        if (args.Length >= 4)
        {
            var heightArg = args[3];
            if (heightArg.IsError)
            {
                return heightArg;
            }

            if (heightArg.Type == CellValueType.Number)
            {
                height = (int)heightArg.NumericValue;
                if (height < 1)
                {
                    return CellValue.Error("#REF!");
                }
            }
        }

        // Extract width (optional, default = 1)
        var width = 1;
        if (args.Length >= 5)
        {
            var widthArg = args[4];
            if (widthArg.IsError)
            {
                return widthArg;
            }

            if (widthArg.Type == CellValueType.Number)
            {
                width = (int)widthArg.NumericValue;
                if (width < 1)
                {
                    return CellValue.Error("#REF!");
                }
            }
        }

        // Calculate the offset position
        var targetRow = baseRow + rowsOffset;
        var targetCol = baseCol + colsOffset;

        // Validate the target is within valid Excel range
        // Excel has max 16384 columns (XFD) and 1048576 rows
        if (targetRow < 1 || targetRow > 1048576 || targetCol < 1 || targetCol > 16384)
        {
            return CellValue.Error("#REF!");
        }

        // Validate that the entire range (if height/width > 1) is within bounds
        if (targetRow + height - 1 > 1048576 || targetCol + width - 1 > 16384)
        {
            return CellValue.Error("#REF!");
        }

        // For Phase 0: Return the value of the single cell at the offset position
        // In a full implementation, this would return a range reference when height/width > 1
        if (context == null)
        {
            return CellValue.Error("#VALUE!");
        }

        var targetReference = GetColumnLetter(targetCol) + targetRow.ToString(CultureInfo.InvariantCulture);
        return context.GetCell(targetReference);
    }

    private static bool TryParseCellReference(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;

        // Remove $ signs for absolute references
        reference = reference.Replace("$", string.Empty);

        // Match cell reference pattern (e.g., A1, B10, AA100)
        var match = Regex.Match(reference, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return false;
        }

        var columnLetters = match.Groups[1].Value;
        var rowPart = match.Groups[2].Value;

        // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
        column = 0;
        for (var i = 0; i < columnLetters.Length; i++)
        {
            column = (column * 26) + (char.ToUpperInvariant(columnLetters[i]) - 'A' + 1);
        }

        if (!int.TryParse(rowPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out row))
        {
            return false;
        }

        return column > 0 && row > 0;
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
