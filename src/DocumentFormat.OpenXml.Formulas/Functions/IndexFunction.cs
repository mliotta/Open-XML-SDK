// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the INDEX function.
/// INDEX(array, row_num, [column_num]) - Returns value at specified position in array.
/// </summary>
public sealed class IndexFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IndexFunction Instance = new();

    private IndexFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "INDEX";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract row_num (always present)
        var rowNumArg = args[args.Length - 1];

        // Check if we have column_num (3 arguments total)
        var hasColumnNum = args.Length >= 3;
        CellValue colNumArg = hasColumnNum ? args[args.Length - 2] : CellValue.Empty;

        // Check for errors in row_num
        if (rowNumArg.IsError)
        {
            return rowNumArg;
        }

        if (rowNumArg.Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rowNum = (int)rowNumArg.NumericValue;

        if (rowNum < 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in column_num if present
        int colNum = 1; // Default to column 1 if not specified
        if (hasColumnNum)
        {
            if (colNumArg.IsError)
            {
                return colNumArg;
            }

            if (colNumArg.Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            colNum = (int)colNumArg.NumericValue;

            if (colNum < 0)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Extract array (everything between start and row_num/col_num arguments)
        var arrayStartIndex = 0;
        var arrayLength = hasColumnNum ? args.Length - 2 : args.Length - 1;

        if (arrayLength == 0)
        {
            return CellValue.Error("#REF!");
        }

        // Check for errors in array
        for (var i = arrayStartIndex; i < arrayStartIndex + arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // If only one cell in array, return it if indices are valid
        if (arrayLength == 1)
        {
            if (rowNum == 1 && colNum == 1)
            {
                return args[arrayStartIndex];
            }
            else
            {
                return CellValue.Error("#REF!");
            }
        }

        // Calculate array dimensions
        // Heuristic: try to find a reasonable shape
        // We prefer shapes close to square (numRows ≈ numCols) as they're more typical in Excel
        var numCols = 0;
        var numRows = 0;
        var bestDiff = int.MaxValue;

        // Special case: if no column number specified, treat as 1D vertical array
        if (!hasColumnNum)
        {
            numRows = arrayLength;
            numCols = 1;
        }
        else
        {
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
        }

        if (numCols == 0 || numRows == 0)
        {
            return CellValue.Error("#REF!");
        }

        // Special case: if row_num is 0, return entire column (not supported - return error)
        if (rowNum == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Special case: if col_num is 0, return entire row (not supported - return error)
        if (colNum == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Validate indices are within bounds
        if (rowNum < 1 || rowNum > numRows || colNum < 1 || colNum > numCols)
        {
            return CellValue.Error("#REF!");
        }

        // Calculate the index in the flattened array
        // Array is in row-major order: row1col1, row1col2, ..., row2col1, row2col2, ...
        var index = arrayStartIndex + ((rowNum - 1) * numCols) + (colNum - 1);

        return args[index];
    }
}
