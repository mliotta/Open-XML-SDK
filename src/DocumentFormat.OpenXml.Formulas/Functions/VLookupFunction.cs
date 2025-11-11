// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the VLOOKUP function.
/// VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) - vertical lookup.
/// </summary>
public sealed class VLookupFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly VLookupFunction Instance = new();

    private VLookupFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "VLOOKUP";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract the last 1 or 2 arguments (col_index and optional range_lookup)
        // Everything between args[0] and these last arguments is the table
        var lastArg = args[args.Length - 1];
        var hasRangeLookup = args.Length >= 4 && (lastArg.Type == CellValueType.Boolean ||
                                                    (lastArg.Type == CellValueType.Number && (lastArg.NumericValue == 0 || lastArg.NumericValue == 1)));

        var colIndexPos = hasRangeLookup ? args.Length - 2 : args.Length - 1;
        var lookupValue = args[0];

        // Validate lookup value
        if (lookupValue.IsError)
        {
            return lookupValue;
        }

        // Validate column index
        if (args[colIndexPos].IsError)
        {
            return args[colIndexPos];
        }

        if (args[colIndexPos].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var colIndex = (int)args[colIndexPos].NumericValue;

        if (colIndex < 1)
        {
            return CellValue.Error("#VALUE!");
        }

        // Determine range lookup mode
        var rangeLookup = true; // Default is approximate match

        if (hasRangeLookup)
        {
            if (lastArg.IsError)
            {
                return lastArg;
            }

            rangeLookup = lastArg.Type switch
            {
                CellValueType.Boolean => lastArg.BoolValue,
                CellValueType.Number => lastArg.NumericValue != 0,
                _ => true,
            };
        }

        // Extract table array (everything between lookup value and col_index)
        var tableStartIndex = 1;
        var tableLength = colIndexPos - 1;

        if (tableLength < colIndex)
        {
            return CellValue.Error("#REF!");
        }

        // Check for errors in table
        for (var i = tableStartIndex; i < tableStartIndex + tableLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Calculate table dimensions
        // Heuristic: try to find a reasonable number of columns
        // The table must have at least col_index columns, and the total length must be divisible by numCols
        var numCols = colIndex;

        // Try to find the smallest valid column count >= colIndex that divides tableLength evenly
        while (numCols <= tableLength && tableLength % numCols != 0)
        {
            numCols++;
        }

        if (tableLength % numCols != 0 || numCols > tableLength)
        {
            return CellValue.Error("#REF!");
        }

        var numRows = tableLength / numCols;

        // VLOOKUP searches the first column of the table
        // Table is in row-major order: row1col1, row1col2, ..., row2col1, row2col2, ...
        for (var row = 0; row < numRows; row++)
        {
            var firstColIndex = tableStartIndex + (row * numCols);
            var firstColValue = args[firstColIndex];

            if (!rangeLookup && ValuesEqual(firstColValue, lookupValue))
            {
                // Exact match found - return value from the specified column
                var resultIndex = firstColIndex + (colIndex - 1);
                return args[resultIndex];
            }
            else if (rangeLookup)
            {
                // Approximate match logic (for sorted data)
                // Find largest value <= lookup value
                if (CompareValues(firstColValue, lookupValue) <= 0)
                {
                    // Check if this is the last matching row
                    var isLast = (row == numRows - 1);
                    if (!isLast)
                    {
                        var nextFirstColIndex = tableStartIndex + ((row + 1) * numCols);
                        var nextFirstColValue = args[nextFirstColIndex];
                        if (CompareValues(nextFirstColValue, lookupValue) > 0)
                        {
                            // This is the last row where first col <= lookup value
                            var resultIndex = firstColIndex + (colIndex - 1);
                            return args[resultIndex];
                        }
                    }
                    else
                    {
                        // Last row
                        var resultIndex = firstColIndex + (colIndex - 1);
                        return args[resultIndex];
                    }
                }
            }
        }

        return CellValue.Error("#N/A");
    }

    private static bool ValuesEqual(CellValue a, CellValue b)
    {
        if (a.Type != b.Type)
        {
            return false;
        }

        return a.Type switch
        {
            CellValueType.Number => System.Math.Abs(a.NumericValue - b.NumericValue) < 1e-10,
            CellValueType.Text => string.Equals(a.StringValue, b.StringValue, StringComparison.OrdinalIgnoreCase),
            CellValueType.Boolean => a.BoolValue == b.BoolValue,
            CellValueType.Empty => true,
            _ => false,
        };
    }

    private static int CompareValues(CellValue a, CellValue b)
    {
        // Compare two values for ordering
        if (a.Type != b.Type)
        {
            // Type mismatch - use type priority: Number < Text < Boolean < Empty
            return a.Type.CompareTo(b.Type);
        }

        return a.Type switch
        {
            CellValueType.Number => a.NumericValue.CompareTo(b.NumericValue),
            CellValueType.Text => string.Compare(a.StringValue, b.StringValue, StringComparison.OrdinalIgnoreCase),
            CellValueType.Boolean => a.BoolValue.CompareTo(b.BoolValue),
            CellValueType.Empty => 0,
            _ => 0,
        };
    }
}
