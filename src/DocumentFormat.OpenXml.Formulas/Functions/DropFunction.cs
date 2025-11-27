// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the DROP function.
/// DROP(array, rows, [columns]) - Drops the first or last N rows or columns from an array.
/// Positive values drop from the start, negative values drop from the end.
/// NOTE: Due to single-value return limitation, only the first element of the result is returned.
/// </summary>
public sealed class DropFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly DropFunction Instance = new();

    private DropFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "DROP";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Parse rows parameter (last argument is always rows)
        if (args[args.Length - 1].IsError)
        {
            return args[args.Length - 1];
        }

        if (args[args.Length - 1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var rows = (int)args[args.Length - 1].NumericValue;

        // Note: Optional columns parameter is not supported in this simplified implementation
        // since we cannot reliably distinguish between array data and the optional parameter
        // when all values are numeric. Only the rows parameter is processed.

        // Determine array length (everything except the last argument which is rows)
        var arrayLength = args.Length - 1;

        if (arrayLength == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // For this simplified implementation, treat the array as a 1D row vector
        var numRows = arrayLength;

        // Validate dimensions - cannot drop more rows than exist
        if (System.Math.Abs(rows) >= numRows)
        {
            return FormulaResult.Error("#CALC!");
        }

        // Determine which element to return (first element after dropping)
        int firstIndex;
        if (rows > 0)
        {
            // Drop from start, return element at index 'rows'
            firstIndex = rows;
        }
        else
        {
            // Drop from end, return first element (index 0)
            firstIndex = 0;
        }

        if (firstIndex >= 0 && firstIndex < arrayLength)
        {
            return args[firstIndex];
        }

        return FormulaResult.Error("#REF!");
    }
}
