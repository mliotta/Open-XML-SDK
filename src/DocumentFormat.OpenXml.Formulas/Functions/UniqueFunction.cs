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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Parse optional parameters from the end
        // Signature: UNIQUE(array, [by_col], [occurs_once])
        // Both parameters are boolean, need to parse in correct order
        var byCol = false;
        var occursOnce = false;
        var arrayLength = args.Length;
        var boolCount = 0;

        // Count boolean args from the end
        for (var i = args.Length - 1; i >= 0 && args[i].Type == FormulaResultType.Boolean; i--)
        {
            boolCount++;
        }

        // Parse based on how many booleans we found
        if (boolCount >= 2)
        {
            // Both parameters present: ..., by_col, occurs_once
            byCol = args[args.Length - 2].BoolValue;
            occursOnce = args[args.Length - 1].BoolValue;
            arrayLength -= 2;
        }
        else if (boolCount == 1)
        {
            // Only one boolean - need to determine if it's by_col or occurs_once
            // Heuristic: if it's TRUE, more likely to be occurs_once (filtering to unique-once)
            // If it's FALSE, more likely to be by_col=FALSE (default behavior, compare rows)
            // For now, assume single boolean is occurs_once (more common use case)
            occursOnce = args[args.Length - 1].BoolValue;
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
        // Prefer column vectors (numCols = 1) for 1D data
        var numCols = 0;
        var numRows = 0;

        // First, check if it can be a single column
        if (arrayLength >= 1)
        {
            numCols = 1;
            numRows = arrayLength;
        }

        // If by_col is true, we're comparing columns, so prefer single row instead
        if (byCol && arrayLength >= 1)
        {
            numCols = arrayLength;
            numRows = 1;
        }

        if (numCols == 0 || numRows == 0)
        {
            return FormulaResult.Error("#REF!");
        }

        if (!byCol)
        {
            // Compare rows for uniqueness - preserve insertion order
            var rowMap = new Dictionary<string, RowInfo>();
            var insertionOrder = new List<string>();

            for (var row = 0; row < numRows; row++)
            {
                var rowValues = new FormulaResult[numCols];
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
                    insertionOrder.Add(key);
                }
            }

            // Filter based on occursOnce, preserving insertion order
            var uniqueRows = new List<FormulaResult[]>();
            foreach (var key in insertionOrder)
            {
                var entry = rowMap[key];
                if (!occursOnce || entry.Count == 1)
                {
                    uniqueRows.Add(entry.Values);
                }
            }

            // If no unique values found
            if (uniqueRows.Count == 0)
            {
                return FormulaResult.Error("#CALC!");
            }

            // Flatten to array
            var resultLength = uniqueRows.Count * numCols;
            var result = new FormulaResult[resultLength];
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
            // Compare columns for uniqueness - preserve insertion order
            var colMap = new Dictionary<string, RowInfo>();
            var insertionOrder = new List<string>();

            for (var col = 0; col < numCols; col++)
            {
                var colValues = new FormulaResult[numRows];
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
                    insertionOrder.Add(key);
                }
            }

            // Filter based on occursOnce, preserving insertion order
            var uniqueCols = new List<FormulaResult[]>();
            foreach (var key in insertionOrder)
            {
                var entry = colMap[key];
                if (!occursOnce || entry.Count == 1)
                {
                    uniqueCols.Add(entry.Values);
                }
            }

            // If no unique values found
            if (uniqueCols.Count == 0)
            {
                return FormulaResult.Error("#CALC!");
            }

            // Flatten to array (reorganize as row-major)
            var resultLength = numRows * uniqueCols.Count;
            var result = new FormulaResult[resultLength];
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

    private static string CreateRowKey(FormulaResult[] values)
    {
        // Create a unique string key from cell values
        var parts = new string[values.Length];
        for (var i = 0; i < values.Length; i++)
        {
            var val = values[i];
            switch (val.Type)
            {
                case FormulaResultType.Number:
                    parts[i] = "N:" + val.NumericValue.ToString();
                    break;
                case FormulaResultType.Text:
                    parts[i] = "T:" + val.StringValue;
                    break;
                case FormulaResultType.Boolean:
                    parts[i] = "B:" + val.BoolValue.ToString();
                    break;
                case FormulaResultType.Empty:
                    parts[i] = "E:";
                    break;
                case FormulaResultType.Error:
                    parts[i] = "ERR:" + val.ErrorValue;
                    break;
                default:
                    parts[i] = "?";
                    break;
            }
        }
        return string.Join("|", parts);
    }

    private sealed class RowInfo
    {
        public FormulaResult[] Values { get; set; } = null!;
        public int Count { get; set; }
    }
}
