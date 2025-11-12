// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COLUMN function.
/// COLUMN([reference]) - Returns the column number of a reference.
/// If no reference is provided, returns the column number of the current cell.
/// </summary>
public sealed class ColumnFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ColumnFunction Instance = new();

    private ColumnFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COLUMN";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            // No reference provided - return current cell's column
            if (context?.CurrentCellReference == null)
            {
                return CellValue.Error("#VALUE!");
            }

            var column = ParseColumnFromReference(context.CurrentCellReference);
            return CellValue.FromNumber(column);
        }

        if (args.Length == 1)
        {
            // Reference provided as a cell value
            var reference = args[0];

            if (reference.IsError)
            {
                return reference;
            }

            if (reference.Type == CellValueType.Text)
            {
                // Try to parse as cell reference
                var column = ParseColumnFromReference(reference.StringValue);
                if (column > 0)
                {
                    return CellValue.FromNumber(column);
                }

                return CellValue.Error("#VALUE!");
            }

            // For array references, we would need to return an array of column numbers
            // For now, return error
            return CellValue.Error("#VALUE!");
        }

        return CellValue.Error("#VALUE!");
    }

    private static int ParseColumnFromReference(string reference)
    {
        // Remove $ signs for absolute references
        reference = reference.Replace("$", string.Empty);

        // Match cell reference pattern (e.g., A1, B10, AA100)
        var match = Regex.Match(reference, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return -1;
        }

        var columnLetters = match.Groups[1].Value;

        // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
        var column = 0;
        for (var i = 0; i < columnLetters.Length; i++)
        {
            column = (column * 26) + (char.ToUpperInvariant(columnLetters[i]) - 'A' + 1);
        }

        return column;
    }
}
