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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Determine argument layout
        // Signature: INDEX(array_elements..., row_num, [col_num])
        // We need at least: 1 array element + row_num = 2 args minimum
        // With column: 1 array element + row_num + col_num = 3 args minimum

        // Strategy: Try to determine if the last 1 or 2 args are position indicators
        // by checking if they're both numbers within a reasonable range for array dimensions
        var lastArg = args[args.Length - 1];
        var secondToLastArg = args.Length >= 3 ? args[args.Length - 2] : FormulaResult.Empty;

        bool hasColumnNum = false;
        FormulaResult rowNumArg;
        FormulaResult colNumArg = FormulaResult.Empty;
        int arrayLength;

        // Try interpretation with 2 position args (row, col)
        // Use heuristic: if last two args are small positive integers, treat as positions
        // "Small" means <= 10 to distinguish from typical array values
        bool canBeTwoPositions = args.Length >= 3 &&
                                secondToLastArg.Type == FormulaResultType.Number &&
                                lastArg.Type == FormulaResultType.Number &&
                                secondToLastArg.NumericValue >= 1 &&
                                secondToLastArg.NumericValue <= 10 && // Small position range
                                lastArg.NumericValue >= 1 &&
                                lastArg.NumericValue <= 10; // Small position range

        if (canBeTwoPositions)
        {
            // Use 2-position interpretation: array..., row_num, col_num
            // Don't validate if positions are valid for array - that happens later
            rowNumArg = secondToLastArg;
            colNumArg = lastArg;
            hasColumnNum = true;
            arrayLength = args.Length - 2;
        }
        else
        {
            // Only one position arg
            rowNumArg = lastArg;
            hasColumnNum = false;
            arrayLength = args.Length - 1;
        }

        // Check for errors in row_num
        if (rowNumArg.IsError)
        {
            return rowNumArg;
        }

        if (rowNumArg.Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var rowNum = (int)rowNumArg.NumericValue;

        if (rowNum < 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in column_num if present
        int colNum = 1; // Default to column 1 if not specified
        if (hasColumnNum)
        {
            if (colNumArg.IsError)
            {
                return colNumArg;
            }

            if (colNumArg.Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            colNum = (int)colNumArg.NumericValue;

            if (colNum < 0)
            {
                return FormulaResult.Error("#VALUE!");
            }
        }

        // Extract array (everything before row_num/col_num arguments)
        var arrayStartIndex = 0;

        if (arrayLength == 0)
        {
            return FormulaResult.Error("#REF!");
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
                return FormulaResult.Error("#REF!");
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
            return FormulaResult.Error("#REF!");
        }

        // Special case: if row_num is 0, return entire column (not supported - return error)
        if (rowNum == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Special case: if col_num is 0, return entire row (not supported - return error)
        if (colNum == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Validate indices are within bounds
        if (rowNum < 1 || rowNum > numRows || colNum < 1 || colNum > numCols)
        {
            return FormulaResult.Error("#REF!");
        }

        // Calculate the index in the flattened array
        // Array is in row-major order: row1col1, row1col2, ..., row2col1, row2col2, ...
        var index = arrayStartIndex + ((rowNum - 1) * numCols) + (colNum - 1);

        return args[index];
    }
}
