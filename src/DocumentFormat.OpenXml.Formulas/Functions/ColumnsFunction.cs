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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
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
            return CellValue.FromNumber(1);
        }

        // For multiple cells, infer the array dimensions using the same heuristic as INDEX/VLOOKUP
        // We prefer shapes close to square (numRows ≈ numCols) as they're more typical in Excel
        var arrayLength = args.Length;
        var numCols = 0;
        var numRows = 0;
        var bestDiff = int.MaxValue;

        // Find the column count that gives the most square-like shape
        for (var testCols = 1; testCols <= arrayLength; testCols++)
        {
            if (arrayLength % testCols == 0)
            {
                var testRows = arrayLength / testCols;
                var diff = System.Math.Abs(testRows - testCols);
                if (diff < bestDiff)
                {
                    numCols = testCols;
                    numRows = testRows;
                    bestDiff = diff;
                }
            }
        }

        if (numCols == 0)
        {
            // Shouldn't happen for valid arrays, but fallback to treating as single column
            return CellValue.FromNumber(1);
        }

        return CellValue.FromNumber(numCols);
    }
}
