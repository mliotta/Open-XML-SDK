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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse rows parameter
        if (args[args.Length - 1].IsError)
        {
            return args[args.Length - 1];
        }

        if (args[args.Length - 1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rows = (int)args[args.Length - 1].NumericValue;

        // Parse optional columns parameter
        var cols = 0;
        var hasColumns = false;
        if (args.Length >= 3 && args[args.Length - 2].Type == CellValueType.Number)
        {
            cols = (int)args[args.Length - 2].NumericValue;
            hasColumns = true;
        }

        // Determine array length
        var arrayLength = hasColumns ? args.Length - 2 : args.Length - 1;

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

        // Calculate array dimensions
        var numCols = hasColumns ? System.Math.Max(1, System.Math.Abs(cols)) : 1;
        var numRows = arrayLength;

        if (hasColumns && arrayLength % numCols == 0)
        {
            numRows = arrayLength / numCols;
        }

        // Validate dimensions
        if (System.Math.Abs(rows) >= numRows || (hasColumns && System.Math.Abs(cols) >= numCols))
        {
            return CellValue.Error("#CALC!");
        }

        // Determine which elements to keep (opposite of TAKE)
        int startRow, endRow;
        if (rows > 0)
        {
            // Drop from start, keep the rest
            startRow = rows;
            endRow = numRows;
        }
        else
        {
            // Drop from end, keep the beginning
            startRow = 0;
            endRow = numRows + rows; // rows is negative
        }

        int startCol, endCol;
        if (hasColumns)
        {
            if (cols > 0)
            {
                // Drop from start
                startCol = cols;
                endCol = numCols;
            }
            else
            {
                // Drop from end
                startCol = 0;
                endCol = numCols + cols; // cols is negative
            }
        }
        else
        {
            startCol = 0;
            endCol = numCols;
        }

        // Return first element of the remaining range
        var firstIndex = startRow * numCols + startCol;
        if (firstIndex >= 0 && firstIndex < arrayLength)
        {
            return args[firstIndex];
        }

        return CellValue.Error("#REF!");
    }
}
