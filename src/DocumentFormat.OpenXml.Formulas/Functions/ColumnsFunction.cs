// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COLUMNS function.
/// COLUMNS(array) - Returns the number of columns in an array or reference.
/// </summary>
public sealed class ColumnsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ColumnsFunction Instance = new();

    private ColumnsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COLUMNS";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in array
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // If single cell, return 1
        if (args.Length == 1)
        {
            return FormulaResult.FromNumber(1);
        }

        // For multiple cells, infer array dimensions
        // Strategy: For composite numbers, prefer taller arrays (fewer columns)
        // For prime numbers, treat as single row (all columns)
        var arrayLength = args.Length;
        var numCols = arrayLength; // Default to single row

        // Find the smallest column count that evenly divides the array length
        for (var testCols = 2; testCols <= System.Math.Sqrt(arrayLength); testCols++)
        {
            if (arrayLength % testCols == 0)
            {
                numCols = testCols;
                break;
            }
        }

        return FormulaResult.FromNumber(numCols);
    }
}
