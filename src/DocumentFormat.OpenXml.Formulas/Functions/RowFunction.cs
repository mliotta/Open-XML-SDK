// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ROW function.
/// ROW([reference]) - Returns the row number of a reference.
/// If no reference is provided, returns the row number of the current cell.
/// </summary>
public sealed class RowFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RowFunction Instance = new();

    private RowFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ROW";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            // No reference provided - return current cell's row
            if (context?.CurrentCellReference == null)
            {
                return CellValue.Error("#VALUE!");
            }

            var row = ParseRowFromReference(context.CurrentCellReference);
            return CellValue.FromNumber(row);
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
                var row = ParseRowFromReference(reference.StringValue);
                if (row > 0)
                {
                    return CellValue.FromNumber(row);
                }

                return CellValue.Error("#VALUE!");
            }

            // For array references, we would need to return an array of row numbers
            // For now, return error
            return CellValue.Error("#VALUE!");
        }

        return CellValue.Error("#VALUE!");
    }

    private static int ParseRowFromReference(string reference)
    {
        // Remove $ signs for absolute references
        reference = reference.Replace("$", string.Empty);

        // Match cell reference pattern (e.g., A1, B10, AA100)
        var match = Regex.Match(reference, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return -1;
        }

        var rowPart = match.Groups[2].Value;

        if (int.TryParse(rowPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row))
        {
            return row;
        }

        return -1;
    }
}
