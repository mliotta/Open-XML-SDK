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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // In this simplified implementation, the last argument is the column number,
        // and everything before it is the array. We cannot reliably detect multiple
        // column numbers when array data is also numeric.
        if (args[args.Length - 1].IsError)
        {
            return args[args.Length - 1];
        }

        if (args[args.Length - 1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var colNum = (int)args[args.Length - 1].NumericValue;

        if (colNum == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var arrayLength = args.Length - 1;

        if (arrayLength == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // For this simplified implementation, treat the array as a 1D row vector.
        // Column 1 = first element, column 2 = second element, etc.
        // Negative column numbers count from the end.
        var numCols = arrayLength;

        // Convert to 0-based index
        int colIndex;
        if (colNum > 0)
        {
            colIndex = colNum - 1;
        }
        else
        {
            // Negative: -1 means last column, -2 means second to last, etc.
            colIndex = numCols + colNum;
        }

        // Validate column index
        if (colIndex < 0 || colIndex >= numCols)
        {
            return FormulaResult.Error("#VALUE!");
        }

        return args[colIndex];
    }
}
