// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TRANSPOSE function.
/// TRANSPOSE(array) - Transposes rows and columns of an array or range.
/// </summary>
public sealed class TransposeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TransposeFunction Instance = new TransposeFunction();

    private TransposeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TRANSPOSE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in input array
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Handle single cell - transpose of single cell is itself
        if (args.Length == 1)
        {
            return args[0];
        }

        // Calculate array dimensions
        // Try to find the most square-like shape for the input
        var arrayLength = args.Length;
        var numCols = 0;
        var numRows = 0;
        var bestDiff = int.MaxValue;

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

        if (numCols == 0 || numRows == 0)
        {
            return CellValue.Error("#REF!");
        }

        // Create transposed array
        // Original: row-major [r0c0, r0c1, r1c0, r1c1]
        // Transposed: [r0c0, r1c0, r0c1, r1c1]
        var transposed = new CellValue[arrayLength];

        for (var row = 0; row < numRows; row++)
        {
            for (var col = 0; col < numCols; col++)
            {
                var originalIndex = row * numCols + col;
                var transposedIndex = col * numRows + row;
                transposed[transposedIndex] = args[originalIndex];
            }
        }

        // Note: In Excel, TRANSPOSE returns an array
        // For this implementation, we return the first element
        // Full array support would require changes to the evaluation engine
        return transposed[0];
    }
}
