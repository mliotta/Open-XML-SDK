// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SORT function.
/// SORT(array, [sort_index], [sort_order], [by_col]) - Sorts the contents of a range or array.
/// sort_index: Column/row number to sort by (default: 1)
/// sort_order: 1 for ascending (default), -1 for descending
/// by_col: FALSE to sort by rows (default), TRUE to sort by columns
/// </summary>
public sealed class SortFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SortFunction Instance = new SortFunction();

    private SortFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SORT";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Parse optional parameters from the end
        var sortIndex = 1;
        var sortOrder = 1; // 1 = ascending, -1 = descending
        var byCol = false;
        var arrayLength = args.Length;

        // Check if last argument is by_col (boolean)
        if (args.Length >= 2 && args[args.Length - 1].Type == FormulaResultType.Boolean)
        {
            byCol = args[args.Length - 1].BoolValue;
            arrayLength--;
        }

        // Check if second-to-last (or last if no by_col) is sort_order
        if (arrayLength >= 2 && args[arrayLength - 1].Type == FormulaResultType.Number)
        {
            var orderValue = args[arrayLength - 1].NumericValue;
            if (orderValue == 1 || orderValue == -1)
            {
                sortOrder = (int)orderValue;
                arrayLength--;
            }
        }

        // Check if third-to-last (or current last) is sort_index
        if (arrayLength >= 2 && args[arrayLength - 1].Type == FormulaResultType.Number)
        {
            sortIndex = (int)args[arrayLength - 1].NumericValue;
            if (sortIndex < 1)
            {
                return FormulaResult.Error("#VALUE!");
            }
            arrayLength--;
        }

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Handle single cell
        if (arrayLength == 1)
        {
            return args[0];
        }

        // Calculate array dimensions
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
            return FormulaResult.Error("#REF!");
        }

        // Validate sort_index is within bounds
        if (byCol && sortIndex > numRows)
        {
            return FormulaResult.Error("#VALUE!");
        }
        else if (!byCol && sortIndex > numCols)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Sort by rows (default behavior)
        if (!byCol)
        {
            // Create list of rows
            var rows = new List<RowData>();
            for (var row = 0; row < numRows; row++)
            {
                var rowValues = new FormulaResult[numCols];
                for (var col = 0; col < numCols; col++)
                {
                    rowValues[col] = args[row * numCols + col];
                }
                rows.Add(new RowData { Index = row, Values = rowValues });
            }

            // Sort rows by the specified column
            var sortIndexLocal = sortIndex;
            var sortOrderLocal = sortOrder;
            rows.Sort((a, b) =>
            {
                var compareResult = CompareValues(a.Values[sortIndexLocal - 1], b.Values[sortIndexLocal - 1]);
                return sortOrderLocal * compareResult;
            });

            // Flatten sorted rows back to array
            var sorted = new FormulaResult[arrayLength];
            for (var i = 0; i < numRows; i++)
            {
                for (var col = 0; col < numCols; col++)
                {
                    sorted[i * numCols + col] = rows[i].Values[col];
                }
            }

            // Return first element (full array support would require engine changes)
            return sorted[0];
        }
        else
        {
            // Sort by columns
            var cols = new List<RowData>();
            for (var col = 0; col < numCols; col++)
            {
                var colValues = new FormulaResult[numRows];
                for (var row = 0; row < numRows; row++)
                {
                    colValues[row] = args[row * numCols + col];
                }
                cols.Add(new RowData { Index = col, Values = colValues });
            }

            // Sort columns by the specified row
            var sortIndexLocal = sortIndex;
            var sortOrderLocal = sortOrder;
            cols.Sort((a, b) =>
            {
                var compareResult = CompareValues(a.Values[sortIndexLocal - 1], b.Values[sortIndexLocal - 1]);
                return sortOrderLocal * compareResult;
            });

            // Flatten sorted columns back to array
            var sorted = new FormulaResult[arrayLength];
            for (var row = 0; row < numRows; row++)
            {
                for (var i = 0; i < numCols; i++)
                {
                    sorted[row * numCols + i] = cols[i].Values[row];
                }
            }

            return sorted[0];
        }
    }

    private static int CompareValues(FormulaResult a, FormulaResult b)
    {
        // Empty values sort last
        if (a.Type == FormulaResultType.Empty && b.Type == FormulaResultType.Empty)
        {
            return 0;
        }
        if (a.Type == FormulaResultType.Empty)
        {
            return 1;
        }
        if (b.Type == FormulaResultType.Empty)
        {
            return -1;
        }

        // Errors sort last
        if (a.IsError && b.IsError)
        {
            return 0;
        }
        if (a.IsError)
        {
            return 1;
        }
        if (b.IsError)
        {
            return -1;
        }

        // Same type comparison
        if (a.Type == b.Type)
        {
            switch (a.Type)
            {
                case FormulaResultType.Number:
                    return a.NumericValue.CompareTo(b.NumericValue);
                case FormulaResultType.Text:
                    return string.Compare(a.StringValue, b.StringValue, StringComparison.OrdinalIgnoreCase);
                case FormulaResultType.Boolean:
                    return a.BoolValue.CompareTo(b.BoolValue);
                default:
                    return 0;
            }
        }

        // Different types: Numbers < Text < Boolean
        return a.Type.CompareTo(b.Type);
    }

    private sealed class RowData
    {
        public int Index { get; set; }
        public FormulaResult[] Values { get; set; } = null!;
    }
}
