// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the INDIRECT function.
/// INDIRECT(ref_text, [a1]) - Returns reference specified by text string.
/// For Phase 0, returns the cell value at the referenced position.
/// </summary>
public sealed class IndirectFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IndirectFunction Instance = new();

    private IndirectFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "INDIRECT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract ref_text (required)
        var refTextArg = args[0];
        if (refTextArg.IsError)
        {
            return refTextArg;
        }

        if (refTextArg.Type != CellValueType.Text)
        {
            return CellValue.Error("#VALUE!");
        }

        var refText = refTextArg.StringValue;

        // Extract a1 (optional, default = TRUE for A1 notation)
        var useA1 = true;
        if (args.Length >= 2)
        {
            var a1Arg = args[1];
            if (a1Arg.IsError)
            {
                return a1Arg;
            }

            if (a1Arg.Type == CellValueType.Boolean)
            {
                useA1 = a1Arg.BoolValue;
            }
            else if (a1Arg.Type == CellValueType.Number)
            {
                useA1 = a1Arg.NumericValue != 0;
            }
        }

        // Parse the reference text based on notation
        string cellReference;
        if (useA1)
        {
            // A1 notation (e.g., "A1", "B10", "$C$5")
            cellReference = ParseA1Reference(refText);
        }
        else
        {
            // R1C1 notation (e.g., "R1C1", "R[1]C[2]")
            cellReference = ParseR1C1Reference(refText, context);
        }

        if (cellReference == null)
        {
            return CellValue.Error("#REF!");
        }

        // For Phase 0: Return the value at the referenced cell
        if (context == null)
        {
            return CellValue.Error("#VALUE!");
        }

        return context.GetCell(cellReference);
    }

    private static string? ParseA1Reference(string refText)
    {
        // Remove any sheet name prefix (e.g., "Sheet1!A1" -> "A1")
        var parts = refText.Split('!');
        var reference = parts.Length > 1 ? parts[1] : refText;

        // Remove quotes around sheet names if present
        reference = reference.Trim('\'');

        // Validate A1 reference pattern
        var match = Regex.Match(reference, @"^(\$?)([A-Z]+)(\$?)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return null;
        }

        return reference;
    }

    private static string? ParseR1C1Reference(string refText, CellContext? context)
    {
        // Remove any sheet name prefix
        var parts = refText.Split('!');
        var reference = parts.Length > 1 ? parts[1] : refText;

        // Parse R1C1 notation: R[offset]C[offset] or R#C#
        // Examples: R1C1 (absolute), R[1]C[2] (relative), R[-1]C[0] (relative)
        var match = Regex.Match(reference, @"^R(\[?-?\d+\]?)C(\[?-?\d+\]?)$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return null;
        }

        var rowPart = match.Groups[1].Value;
        var colPart = match.Groups[2].Value;

        int row, col;

        // Parse row
        if (rowPart.StartsWith("[") && rowPart.EndsWith("]"))
        {
            // Relative row reference: R[offset]
            var offset = int.Parse(rowPart.Substring(1, rowPart.Length - 2), CultureInfo.InvariantCulture);

            // For relative references, we need the current cell context
            if (context?.CurrentCellReference == null)
            {
                return null;
            }

            if (!TryParseCellReference(context.CurrentCellReference, out _, out var currentRow))
            {
                return null;
            }

            row = currentRow + offset;
        }
        else
        {
            // Absolute row reference: R#
            row = int.Parse(rowPart, CultureInfo.InvariantCulture);
        }

        // Parse column
        if (colPart.StartsWith("[") && colPart.EndsWith("]"))
        {
            // Relative column reference: C[offset]
            var offset = int.Parse(colPart.Substring(1, colPart.Length - 2), CultureInfo.InvariantCulture);

            // For relative references, we need the current cell context
            if (context?.CurrentCellReference == null)
            {
                return null;
            }

            if (!TryParseCellReference(context.CurrentCellReference, out var currentCol, out _))
            {
                return null;
            }

            col = currentCol + offset;
        }
        else
        {
            // Absolute column reference: C#
            col = int.Parse(colPart, CultureInfo.InvariantCulture);
        }

        // Validate range
        if (row < 1 || row > 1048576 || col < 1 || col > 16384)
        {
            return null;
        }

        // Convert to A1 notation
        return GetColumnLetter(col) + row.ToString(CultureInfo.InvariantCulture);
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
