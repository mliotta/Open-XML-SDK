// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SORTBY function.
/// SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...) - Sorts an array based on values in another array.
/// sort_order: 1 for ascending (default), -1 for descending
/// NOTE: Due to single-value return limitation, only the first element of the sorted array is returned.
/// </summary>
public sealed class SortByFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SortByFunction Instance = new();

    private SortByFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SORTBY";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // In this simplified implementation, we assume the args are structured as:
        // [array elements..., by_array elements...]
        // where both arrays have the same length (args.Length / 2).
        // The optional sort_order parameter is not supported in this simplified version.
        if (args.Length % 2 != 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var arrayLength = args.Length / 2;

        if (arrayLength == 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in both arrays
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Extract the by_array (second half of args)
        var byArray = new FormulaResult[arrayLength];
        for (var i = 0; i < arrayLength; i++)
        {
            byArray[i] = args[arrayLength + i];
        }

        // Create indexed list
        var indexed = new List<IndexedValue>();
        for (var i = 0; i < arrayLength; i++)
        {
            indexed.Add(new IndexedValue { Index = i, Value = args[i] });
        }

        // Sort by the by_array values (ascending by default)
        indexed.Sort((a, b) => CompareValues(byArray[a.Index], byArray[b.Index]));

        // Return first element of sorted array
        return indexed[0].Value;
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

    private sealed class IndexedValue
    {
        public int Index { get; set; }
        public FormulaResult Value { get; set; }
    }
}
