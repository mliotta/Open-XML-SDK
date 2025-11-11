// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MATCH function.
/// MATCH(lookup_value, lookup_array, [match_type]) - Returns position of value in array.
/// match_type: 1 (default, largest value ≤ lookup_value, array sorted ascending),
///            0 (exact match),
///            -1 (smallest value ≥ lookup_value, array sorted descending).
/// </summary>
public sealed class MatchFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MatchFunction Instance = new();

    private MatchFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MATCH";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // Extract lookup_value (first argument)
        var lookupValue = args[0];

        // Validate lookup value
        if (lookupValue.IsError)
        {
            return lookupValue;
        }

        // Determine if we have match_type (last argument)
        var lastArg = args[args.Length - 1];
        var hasMatchType = args.Length >= 3 && lastArg.Type == CellValueType.Number;

        // Default match type is 1
        var matchType = 1;

        if (hasMatchType)
        {
            if (lastArg.IsError)
            {
                return lastArg;
            }

            matchType = (int)lastArg.NumericValue;

            // Validate match type (-1, 0, or 1)
            if (matchType != -1 && matchType != 0 && matchType != 1)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Extract lookup_array (everything between lookup_value and optional match_type)
        var arrayStartIndex = 1;
        var arrayLength = hasMatchType ? args.Length - 2 : args.Length - 1;

        if (arrayLength == 0)
        {
            return CellValue.Error("#N/A");
        }

        // Check for errors in array
        for (var i = arrayStartIndex; i < arrayStartIndex + arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Search based on match type
        if (matchType == 0)
        {
            // Exact match
            for (var i = 0; i < arrayLength; i++)
            {
                var arrayValue = args[arrayStartIndex + i];
                if (ValuesEqual(arrayValue, lookupValue))
                {
                    // Return 1-based position
                    return CellValue.FromNumber(i + 1);
                }
            }

            // No match found
            return CellValue.Error("#N/A");
        }
        else if (matchType == 1)
        {
            // Find largest value <= lookup_value (assumes sorted ascending)
            int lastMatchIndex = -1;

            for (var i = 0; i < arrayLength; i++)
            {
                var arrayValue = args[arrayStartIndex + i];
                var comparison = CompareValues(arrayValue, lookupValue);

                if (comparison <= 0)
                {
                    // This value is <= lookup value
                    lastMatchIndex = i;
                }
                else
                {
                    // We've gone past the lookup value, stop searching
                    break;
                }
            }

            if (lastMatchIndex >= 0)
            {
                // Return 1-based position
                return CellValue.FromNumber(lastMatchIndex + 1);
            }
            else
            {
                // No value <= lookup_value found
                return CellValue.Error("#N/A");
            }
        }
        else // matchType == -1
        {
            // Find smallest value >= lookup_value (assumes sorted descending)
            int lastMatchIndex = -1;

            for (var i = 0; i < arrayLength; i++)
            {
                var arrayValue = args[arrayStartIndex + i];
                var comparison = CompareValues(arrayValue, lookupValue);

                if (comparison >= 0)
                {
                    // This value is >= lookup value
                    lastMatchIndex = i;
                }
                else
                {
                    // We've gone past the lookup value, stop searching
                    break;
                }
            }

            if (lastMatchIndex >= 0)
            {
                // Return 1-based position
                return CellValue.FromNumber(lastMatchIndex + 1);
            }
            else
            {
                // No value >= lookup_value found
                return CellValue.Error("#N/A");
            }
        }
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
