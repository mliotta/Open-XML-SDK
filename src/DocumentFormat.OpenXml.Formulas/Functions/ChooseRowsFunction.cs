// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // In this simplified implementation, the last argument is the row number,
        // and everything before it is the array. We cannot reliably detect multiple
        // row numbers when array data is also numeric.
        if (args[args.Length - 1].IsError)
        {
            return args[args.Length - 1];
        }

        if (args[args.Length - 1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var rowNum = (int)args[args.Length - 1].NumericValue;

        if (rowNum == 0)
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
        // Row 1 = first element, row 2 = second element, etc.
        // Negative row numbers count from the end.
        var numRows = arrayLength;

        // Convert to 0-based index
        int rowIndex;
        if (rowNum > 0)
        {
            rowIndex = rowNum - 1;
        }
        else
        {
            // Negative: -1 means last row, -2 means second to last, etc.
            rowIndex = numRows + rowNum;
        }

        // Validate row index
        if (rowIndex < 0 || rowIndex >= numRows)
        {
            return FormulaResult.Error("#VALUE!");
        }

        return args[rowIndex];
    }
}
