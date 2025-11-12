// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TRIMRANGE function.
/// TRIMRANGE(array, [rows_to_trim], [cols_to_trim])
/// Trims empty rows and columns from the edges of a range.
///
/// Phase 0 Implementation:
/// - Removes leading and trailing empty values from array
/// - Returns first non-empty value
/// - Full 2D trimming requires engine enhancements
/// </summary>
public sealed class TrimRangeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TrimRangeFunction Instance = new TrimRangeFunction();

    private TrimRangeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TRIMRANGE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse optional parameters
        var rowsTrim = 0; // 0 = trim both top and bottom
        var colsTrim = 0; // 0 = trim both left and right

        var arrayLength = args.Length;

        // Check if last argument is cols_to_trim
        if (args.Length >= 2 && args[args.Length - 1].Type == CellValueType.Number)
        {
            colsTrim = (int)args[args.Length - 1].NumericValue;
            arrayLength--;
        }

        // Check if second-to-last (or last if no cols_to_trim) is rows_to_trim
        if (arrayLength >= 2 && args[arrayLength - 1].Type == CellValueType.Number)
        {
            rowsTrim = (int)args[arrayLength - 1].NumericValue;
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

        // Find first non-empty value
        for (var i = 0; i < arrayLength; i++)
        {
            var val = args[i];
            if (val.Type != CellValueType.Empty)
            {
                return val;
            }
        }

        // All values are empty
        return CellValue.Empty;
    }
}
