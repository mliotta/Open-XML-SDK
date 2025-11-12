// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
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
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // First argument is the array to sort
        var arrayStart = 0;
        var arrayLength = 0;

        // Find where the array ends (first argument)
        for (var i = 0; i < args.Length; i++)
        {
            if (i == 0 || args[i].IsError || args[i].Type != CellValueType.Empty)
            {
                arrayLength++;
            }
            else
            {
                break;
            }
        }

        if (arrayLength == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in array
        for (var i = arrayStart; i < arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Parse sort criteria (by_array, sort_order pairs)
        var sortCriteria = new List<SortCriterion>();
        var idx = arrayLength;

        while (idx < args.Length)
        {
            // Get by_array (same length as main array)
            if (idx >= args.Length)
            {
                break;
            }

            var byArray = new CellValue[arrayLength];
            for (var i = 0; i < arrayLength && idx < args.Length; i++, idx++)
            {
                byArray[i] = args[idx];
                if (args[idx].IsError)
                {
                    return args[idx];
                }
            }

            // Get optional sort_order
            var sortOrder = 1;
            if (idx < args.Length && args[idx].Type == CellValueType.Number)
            {
                var orderValue = args[idx].NumericValue;
                if (orderValue == 1 || orderValue == -1)
                {
                    sortOrder = (int)orderValue;
                    idx++;
                }
            }

            sortCriteria.Add(new SortCriterion { ByArray = byArray, SortOrder = sortOrder });
        }

        if (sortCriteria.Count == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Create indexed list
        var indexed = new List<IndexedValue>();
        for (var i = 0; i < arrayLength; i++)
        {
            indexed.Add(new IndexedValue { Index = i, Value = args[i] });
        }

        // Sort using multiple criteria
        indexed.Sort((a, b) =>
        {
            foreach (var criterion in sortCriteria)
            {
                var compareResult = CompareValues(criterion.ByArray[a.Index], criterion.ByArray[b.Index]);
                if (compareResult != 0)
                {
                    return criterion.SortOrder * compareResult;
                }
            }
            return 0;
        });

        // Return first element of sorted array
        return indexed[0].Value;
    }

    private static int CompareValues(CellValue a, CellValue b)
    {
        // Empty values sort last
        if (a.Type == CellValueType.Empty && b.Type == CellValueType.Empty)
        {
            return 0;
        }
        if (a.Type == CellValueType.Empty)
        {
            return 1;
        }
        if (b.Type == CellValueType.Empty)
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
                case CellValueType.Number:
                    return a.NumericValue.CompareTo(b.NumericValue);
                case CellValueType.Text:
                    return string.Compare(a.StringValue, b.StringValue, StringComparison.OrdinalIgnoreCase);
                case CellValueType.Boolean:
                    return a.BoolValue.CompareTo(b.BoolValue);
                default:
                    return 0;
            }
        }

        // Different types: Numbers < Text < Boolean
        return a.Type.CompareTo(b.Type);
    }

    private class SortCriterion
    {
        public CellValue[] ByArray { get; set; } = new CellValue[0];
        public int SortOrder { get; set; }
    }

    private class IndexedValue
    {
        public int Index { get; set; }
        public CellValue Value { get; set; }
    }
}
