// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FILTER function.
/// FILTER(array, include, [if_empty]) - Filters a range based on criteria.
/// array: The array or range to filter
/// include: A boolean array indicating which rows to include
/// if_empty: Value to return if no items meet criteria (default: #CALC! error)
/// </summary>
public sealed class FilterFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FilterFunction Instance = new FilterFunction();

    private FilterFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FILTER";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse if_empty value if provided
        CellValue ifEmptyValue = CellValue.Error("#CALC!");
        var hasIfEmpty = false;

        // Check if last argument is if_empty
        if (args.Length >= 3 && args[args.Length - 1].Type != CellValueType.Boolean)
        {
            ifEmptyValue = args[args.Length - 1];
            hasIfEmpty = true;
        }

        // Split args into array and include sections
        // We need to determine where array ends and include begins
        // Heuristic: split arguments roughly in half
        var totalArgs = hasIfEmpty ? args.Length - 1 : args.Length;
        var arrayLength = totalArgs / 2;
        var includeLength = totalArgs - arrayLength;

        // Check for errors in array
        for (var i = 0; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Check for errors in include array
        for (var i = arrayLength; i < arrayLength + includeLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
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
            return CellValue.Error("#REF!");
        }

        // Validate include array matches array row count
        if (includeLength != numRows)
        {
            return CellValue.Error("#VALUE!");
        }

        // Filter rows based on include criteria
        var filteredRows = new List<CellValue[]>();

        for (var row = 0; row < numRows; row++)
        {
            var includeValue = args[arrayLength + row];

            // Convert include value to boolean
            var shouldInclude = false;
            if (includeValue.Type == CellValueType.Boolean)
            {
                shouldInclude = includeValue.BoolValue;
            }
            else if (includeValue.Type == CellValueType.Number)
            {
                shouldInclude = includeValue.NumericValue != 0;
            }

            if (shouldInclude)
            {
                var rowValues = new CellValue[numCols];
                for (var col = 0; col < numCols; col++)
                {
                    rowValues[col] = args[row * numCols + col];
                }
                filteredRows.Add(rowValues);
            }
        }

        // If no rows match, return if_empty value
        if (filteredRows.Count == 0)
        {
            return ifEmptyValue;
        }

        // Flatten filtered rows to array
        var resultLength = filteredRows.Count * numCols;
        var result = new CellValue[resultLength];
        for (var i = 0; i < filteredRows.Count; i++)
        {
            for (var col = 0; col < numCols; col++)
            {
                result[i * numCols + col] = filteredRows[i][col];
            }
        }

        // Return first element (full array support would require engine changes)
        return result[0];
    }
}
