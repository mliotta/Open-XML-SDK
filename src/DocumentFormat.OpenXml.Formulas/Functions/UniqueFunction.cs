// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the UNIQUE function.
/// UNIQUE(array, [by_col], [occurs_once]) - Returns unique values from a range or array.
/// array: The array or range to filter
/// by_col: FALSE to compare rows (default), TRUE to compare columns
/// occurs_once: FALSE to return all unique values (default), TRUE to return values that occur exactly once
/// </summary>
public sealed class UniqueFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly UniqueFunction Instance = new UniqueFunction();

    private UniqueFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "UNIQUE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse optional parameters from the end
        var byCol = false;
        var occursOnce = false;
        var arrayLength = args.Length;

        // Check if last argument is occurs_once (boolean)
        if (args.Length >= 2 && args[args.Length - 1].Type == CellValueType.Boolean)
        {
            occursOnce = args[args.Length - 1].BoolValue;
            arrayLength--;
        }

        // Check if second-to-last (or last if no occurs_once) is by_col
        if (arrayLength >= 2 && args[arrayLength - 1].Type == CellValueType.Boolean)
        {
            byCol = args[arrayLength - 1].BoolValue;
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
            return CellValue.Error("#REF!");
        }

        if (!byCol)
        {
            // Compare rows for uniqueness
            var rowMap = new Dictionary<string, RowInfo>();

            for (var row = 0; row < numRows; row++)
            {
                var rowValues = new CellValue[numCols];
                for (var col = 0; col < numCols; col++)
                {
                    rowValues[col] = args[row * numCols + col];
                }

                // Create key from row values
                var key = CreateRowKey(rowValues);

                if (rowMap.ContainsKey(key))
                {
                    var existing = rowMap[key];
                    rowMap[key] = new RowInfo { Values = existing.Values, Count = existing.Count + 1 };
                }
                else
                {
                    rowMap[key] = new RowInfo { Values = rowValues, Count = 1 };
                }
            }

            // Filter based on occursOnce
            var uniqueRows = new List<CellValue[]>();
            foreach (var entry in rowMap.Values)
            {
                if (!occursOnce || entry.Count == 1)
                {
                    uniqueRows.Add(entry.Values);
                }
            }

            // If no unique values found
            if (uniqueRows.Count == 0)
            {
                return CellValue.Error("#CALC!");
            }

            // Flatten to array
            var resultLength = uniqueRows.Count * numCols;
            var result = new CellValue[resultLength];
            for (var i = 0; i < uniqueRows.Count; i++)
            {
                for (var col = 0; col < numCols; col++)
                {
                    result[i * numCols + col] = uniqueRows[i][col];
                }
            }

            return result[0];
        }
        else
        {
            // Compare columns for uniqueness
            var colMap = new Dictionary<string, RowInfo>();

            for (var col = 0; col < numCols; col++)
            {
                var colValues = new CellValue[numRows];
                for (var row = 0; row < numRows; row++)
                {
                    colValues[row] = args[row * numCols + col];
                }

                // Create key from column values
                var key = CreateRowKey(colValues);

                if (colMap.ContainsKey(key))
                {
                    var existing = colMap[key];
                    colMap[key] = new RowInfo { Values = existing.Values, Count = existing.Count + 1 };
                }
                else
                {
                    colMap[key] = new RowInfo { Values = colValues, Count = 1 };
                }
            }

            // Filter based on occursOnce
            var uniqueCols = new List<CellValue[]>();
            foreach (var entry in colMap.Values)
            {
                if (!occursOnce || entry.Count == 1)
                {
                    uniqueCols.Add(entry.Values);
                }
            }

            // If no unique values found
            if (uniqueCols.Count == 0)
            {
                return CellValue.Error("#CALC!");
            }

            // Flatten to array (reorganize as row-major)
            var resultLength = numRows * uniqueCols.Count;
            var result = new CellValue[resultLength];
            for (var row = 0; row < numRows; row++)
            {
                for (var i = 0; i < uniqueCols.Count; i++)
                {
                    result[row * uniqueCols.Count + i] = uniqueCols[i][row];
                }
            }

            return result[0];
        }
    }

    private static string CreateRowKey(CellValue[] values)
    {
        // Create a unique string key from cell values
        var parts = new string[values.Length];
        for (var i = 0; i < values.Length; i++)
        {
            var val = values[i];
            switch (val.Type)
            {
                case CellValueType.Number:
                    parts[i] = "N:" + val.NumericValue.ToString();
                    break;
                case CellValueType.Text:
                    parts[i] = "T:" + val.StringValue;
                    break;
                case CellValueType.Boolean:
                    parts[i] = "B:" + val.BoolValue.ToString();
                    break;
                case CellValueType.Empty:
                    parts[i] = "E:";
                    break;
                case CellValueType.Error:
                    parts[i] = "ERR:" + val.ErrorValue;
                    break;
                default:
                    parts[i] = "?";
                    break;
            }
        }
        return string.Join("|", parts);
    }

    private class RowInfo
    {
        public CellValue[] Values { get; set; }
        public int Count { get; set; }
    }
}
