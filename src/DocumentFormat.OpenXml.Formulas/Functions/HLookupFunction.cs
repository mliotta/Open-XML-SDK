// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the HLOOKUP function.
/// HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup]) - horizontal lookup.
/// </summary>
public sealed class HLookupFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly HLookupFunction Instance = new();

    private HLookupFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "HLOOKUP";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract the last 1 or 2 arguments (row_index and optional range_lookup)
        // Everything between args[0] and these last arguments is the table
        var lastArg = args[args.Length - 1];
        var hasRangeLookup = args.Length >= 4 && (lastArg.Type == CellValueType.Boolean ||
                                                    (lastArg.Type == CellValueType.Number && (lastArg.NumericValue == 0 || lastArg.NumericValue == 1)));

        var rowIndexPos = hasRangeLookup ? args.Length - 2 : args.Length - 1;
        var lookupValue = args[0];

        // Validate lookup value
        if (lookupValue.IsError)
        {
            return lookupValue;
        }

        // Validate row index
        if (args[rowIndexPos].IsError)
        {
            return args[rowIndexPos];
        }

        if (args[rowIndexPos].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var rowIndex = (int)args[rowIndexPos].NumericValue;

        if (rowIndex < 1)
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

        // Extract table array (everything between lookup value and row_index)
        var tableStartIndex = 1;
        var tableLength = rowIndexPos - 1;

        if (tableLength < rowIndex)
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
        // Heuristic: try to find a reasonable number of rows
        // The table must have at least row_index rows, and the total length must be divisible by numRows
        var numRows = rowIndex;

        // Try to find the smallest valid row count >= rowIndex that divides tableLength evenly
        while (numRows <= tableLength && tableLength % numRows != 0)
        {
            numRows++;
        }

        if (tableLength % numRows != 0 || numRows > tableLength)
        {
            return CellValue.Error("#REF!");
        }

        var numCols = tableLength / numRows;

        // HLOOKUP searches the first row of the table
        for (var col = 0; col < numCols; col++)
        {
            var firstRowIndex = tableStartIndex + col;
            var firstRowValue = args[firstRowIndex];

            if (!rangeLookup && ValuesEqual(firstRowValue, lookupValue))
            {
                // Exact match found - return value from the specified row
                var resultIndex = tableStartIndex + ((rowIndex - 1) * numCols) + col;
                return args[resultIndex];
            }
            else if (rangeLookup)
            {
                // Approximate match logic (for sorted data)
                // Find largest value <= lookup value
                if (CompareValues(firstRowValue, lookupValue) <= 0)
                {
                    // Check if this is the last matching column
                    var isLast = (col == numCols - 1);
                    if (!isLast)
                    {
                        var nextFirstRowIndex = tableStartIndex + col + 1;
                        var nextFirstRowValue = args[nextFirstRowIndex];
                        if (CompareValues(nextFirstRowValue, lookupValue) > 0)
                        {
                            // This is the last column where first row <= lookup value
                            var resultIndex = tableStartIndex + ((rowIndex - 1) * numCols) + col;
                            return args[resultIndex];
                        }
                    }
                    else
                    {
                        // Last column
                        var resultIndex = tableStartIndex + ((rowIndex - 1) * numCols) + col;
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
