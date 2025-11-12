// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TOCOL function.
/// TOCOL(array, [ignore], [scan_by_column]) - Converts an array to a single column.
/// NOTE: Due to single-value return limitation, only the first element is returned.
/// </summary>
public sealed class ToColFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ToColFunction Instance = new();

    private ToColFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TOCOL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse optional ignore parameter
        var ignore = 0;
        var hasIgnore = args.Length >= 2 && args[args.Length - 1].Type == CellValueType.Number;
        if (hasIgnore)
        {
            ignore = (int)args[args.Length - 1].NumericValue;
        }

        // Parse optional scan_by_column parameter
        var hasScanByCol = args.Length >= 3 && args[args.Length - 2].Type == CellValueType.Boolean;

        // Determine array length
        var arrayLength = args.Length;
        if (hasIgnore)
        {
            arrayLength--;
        }
        if (hasScanByCol)
        {
            arrayLength--;
        }

        if (arrayLength == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Return first non-ignored element
        // ignore: 0 = keep all, 1 = ignore blanks, 2 = ignore errors, 3 = ignore blanks and errors
        for (var i = 0; i < arrayLength; i++)
        {
            var shouldIgnore = false;

            if ((ignore & 1) != 0 && args[i].Type == CellValueType.Empty)
            {
                shouldIgnore = true;
            }

            if ((ignore & 2) != 0 && args[i].IsError)
            {
                shouldIgnore = true;
            }

            if (!shouldIgnore)
            {
                return args[i];
            }
        }

        // All elements ignored
        return CellValue.Error("#CALC!");
    }
}
