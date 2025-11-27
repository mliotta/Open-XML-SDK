// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TAKE function.
/// TAKE(array, rows, [columns]) - Returns the first or last N rows or columns from an array.
/// Positive values take from the start, negative values take from the end.
/// NOTE: Due to single-value return limitation, only the first element of the result is returned.
/// </summary>
public sealed class TakeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TakeFunction Instance = new();

    private TakeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TAKE";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // In this simplified implementation, the last argument is always the rows parameter,
        // and everything before it is the array. We cannot reliably detect the optional
        // columns parameter when array data is also numeric.
        if (args[args.Length - 1].IsError)
        {
            return args[args.Length - 1];
        }

        if (args[args.Length - 1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var rows = (int)args[args.Length - 1].NumericValue;

        // Determine array length (everything except the last argument which is rows)
        var arrayLength = args.Length - 1;

        if (arrayLength == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (rows == 0)
        {
            return FormulaResult.Error("#CALC!");
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

        // Validate dimensions - cannot take more rows than exist
        if (System.Math.Abs(rows) > numRows)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Determine which element to return (first element of taken range)
        int firstIndex;
        if (rows > 0)
        {
            // Take from start, return first element (index 0)
            firstIndex = 0;
        }
        else
        {
            // Take from end, return element at index (numRows + rows)
            // For example: array [10, 20, 30] with rows = -2
            // Result should be [20, 30], first element is at index 1
            firstIndex = numRows + rows;
        }

        if (firstIndex >= 0 && firstIndex < arrayLength)
        {
            return args[firstIndex];
        }

        return FormulaResult.Error("#REF!");
    }
}
