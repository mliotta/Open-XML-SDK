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

        if (rows == 0 || (hasColumns && cols == 0))
        {
            return CellValue.Error("#CALC!");
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
        if (System.Math.Abs(rows) > numRows || (hasColumns && System.Math.Abs(cols) > numCols))
        {
            return CellValue.Error("#VALUE!");
        }

        // Determine which elements to take
        int startRow, endRow;
        if (rows > 0)
        {
            startRow = 0;
            endRow = rows;
        }
        else
        {
            startRow = numRows + rows; // rows is negative
            endRow = numRows;
        }

        int startCol, endCol;
        if (hasColumns)
        {
            if (cols > 0)
            {
                startCol = 0;
                endCol = cols;
            }
            else
            {
                startCol = numCols + cols; // cols is negative
                endCol = numCols;
            }
        }
        else
        {
            startCol = 0;
            endCol = numCols;
        }

        // Return first element of the taken range
        var firstIndex = startRow * numCols + startCol;
        if (firstIndex >= 0 && firstIndex < arrayLength)
        {
            return args[firstIndex];
        }

        return CellValue.Error("#REF!");
    }
}
