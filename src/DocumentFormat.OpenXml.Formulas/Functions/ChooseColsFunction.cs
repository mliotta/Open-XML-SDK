// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CHOOSECOLS function.
/// CHOOSECOLS(array, col_num1, [col_num2], ...) - Returns the specified columns from an array.
/// NOTE: Due to single-value return limitation, only the first element of the result is returned.
/// </summary>
public sealed class ChooseColsFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ChooseColsFunction Instance = new();

    private ChooseColsFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CHOOSECOLS";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse column numbers from the end
        var colCount = 0;
        for (var i = args.Length - 1; i >= 1; i--)
        {
            if (args[i].Type == CellValueType.Number)
            {
                colCount++;
            }
            else
            {
                break;
            }
        }

        if (colCount == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var arrayLength = args.Length - colCount;

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Get the first column number
        var firstColNum = (int)args[arrayLength].NumericValue;

        if (firstColNum == 0)
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

        // Validate column number
        var actualColNum = firstColNum > 0 ? firstColNum : numCols + firstColNum + 1;
        if (actualColNum < 1 || actualColNum > numCols)
        {
            return CellValue.Error("#VALUE!");
        }

        // Return first element of the first chosen column
        var firstIndex = actualColNum - 1;
        if (firstIndex >= 0 && firstIndex < arrayLength)
        {
            return args[firstIndex];
        }

        return CellValue.Error("#REF!");
    }
}
