// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHOOSEROWS function.
/// CHOOSEROWS(array, row_num1, [row_num2], ...) - Returns the specified rows from an array.
/// NOTE: Due to single-value return limitation, only the first element of the result is returned.
/// </summary>
public sealed class ChooseRowsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChooseRowsFunction Instance = new();

    private ChooseRowsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHOOSEROWS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse row numbers from the end
        var rowCount = 0;
        for (var i = args.Length - 1; i >= 1; i--)
        {
            if (args[i].Type == CellValueType.Number)
            {
                rowCount++;
            }
            else
            {
                break;
            }
        }

        if (rowCount == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var arrayLength = args.Length - rowCount;

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Get the first row number
        var firstRowNum = (int)args[arrayLength].NumericValue;

        if (firstRowNum == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Calculate array dimensions (assume square-ish array)
        var numCols = 1;
        var numRows = arrayLength;

        for (var testCols = 1; testCols <= arrayLength; testCols++)
        {
            if (arrayLength % testCols == 0)
            {
                var testRows = arrayLength / testCols;
                var diff = System.Math.Abs(testRows - testCols);
                if (diff <= System.Math.Abs(numRows - numCols))
                {
                    numCols = testCols;
                    numRows = testRows;
                }
            }
        }

        // Validate row number
        var actualRowNum = firstRowNum > 0 ? firstRowNum : numRows + firstRowNum + 1;
        if (actualRowNum < 1 || actualRowNum > numRows)
        {
            return CellValue.Error("#VALUE!");
        }

        // Return first element of the first chosen row
        var firstIndex = (actualRowNum - 1) * numCols;
        if (firstIndex >= 0 && firstIndex < arrayLength)
        {
            return args[firstIndex];
        }

        return CellValue.Error("#REF!");
    }
}
