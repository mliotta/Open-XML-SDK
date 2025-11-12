// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the XMATCH function.
/// XMATCH(lookup_value, lookup_array, [match_mode], [search_mode]).
/// Modern replacement for MATCH with more options.
/// match_mode: 0 (exact match, default), -1 (exact or next smaller), 1 (exact or next larger), 2 (wildcard).
/// search_mode: 1 (search first to last, default), -1 (search last to first), 2 (binary search ascending), -2 (binary search descending).
/// </summary>
public sealed class XMatchFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly XMatchFunction Instance = new();

    private XMatchFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "XMATCH";

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

        // Determine if we have optional parameters (match_mode and search_mode)
        var lastArg = args[args.Length - 1];
        var secondToLastArg = args.Length >= 3 ? args[args.Length - 2] : CellValue.Empty;

        var hasSearchMode = args.Length >= 4 && lastArg.Type == CellValueType.Number;
        var hasMatchMode = args.Length >= 3 && secondToLastArg.Type == CellValueType.Number;

        // Default values
        var matchMode = 0;
        var searchMode = 1;

        // Extract optional parameters
        if (hasSearchMode)
        {
            if (lastArg.IsError)
            {
                return lastArg;
            }

            searchMode = (int)lastArg.NumericValue;
            if (searchMode < -2 || searchMode > 2 || searchMode == 0)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        if (hasMatchMode)
        {
            if (secondToLastArg.IsError)
            {
                return secondToLastArg;
            }

            matchMode = (int)secondToLastArg.NumericValue;
            if (matchMode < -1 || matchMode > 2)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Extract lookup_array (everything between lookup_value and optional parameters)
        var arrayStartIndex = 1;
        var optionalParamsCount = (hasSearchMode ? 1 : 0) + (hasMatchMode ? 1 : 0);
        var arrayLength = args.Length - 1 - optionalParamsCount;

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

        // Perform search based on match_mode and search_mode
        var matchIndex = -1;

        if (matchMode == 0) // Exact match
        {
            matchIndex = FindExactMatch(args, arrayStartIndex, arrayLength, lookupValue, searchMode);
        }
        else if (matchMode == -1) // Exact match or next smaller
        {
            matchIndex = FindExactOrNextSmaller(args, arrayStartIndex, arrayLength, lookupValue, searchMode);
        }
        else if (matchMode == 1) // Exact match or next larger
        {
            matchIndex = FindExactOrNextLarger(args, arrayStartIndex, arrayLength, lookupValue, searchMode);
        }
        else if (matchMode == 2) // Wildcard match
        {
            matchIndex = FindWildcardMatch(args, arrayStartIndex, arrayLength, lookupValue, searchMode);
        }

        if (matchIndex >= 0)
        {
            // Return 1-based position
            return CellValue.FromNumber(matchIndex + 1);
        }

        // No match found
        return CellValue.Error("#N/A");
    }

    private static int FindExactMatch(CellValue[] args, int startIndex, int length, CellValue lookupValue, int searchMode)
    {
        if (searchMode == 1) // First to last
        {
            for (var i = 0; i < length; i++)
            {
                if (ValuesEqual(args[startIndex + i], lookupValue))
                {
                    return i;
                }
            }
        }
        else if (searchMode == -1) // Last to first
        {
            for (var i = length - 1; i >= 0; i--)
            {
                if (ValuesEqual(args[startIndex + i], lookupValue))
                {
                    return i;
                }
            }
        }
        else if (searchMode == 2) // Binary search ascending
        {
            return BinarySearch(args, startIndex, length, lookupValue, true);
        }
        else if (searchMode == -2) // Binary search descending
        {
            return BinarySearch(args, startIndex, length, lookupValue, false);
        }

        return -1;
    }

    private static int FindExactOrNextSmaller(CellValue[] args, int startIndex, int length, CellValue lookupValue, int searchMode)
    {
        var lastMatch = -1;

        if (searchMode == 1 || searchMode == 2) // Forward search
        {
            for (var i = 0; i < length; i++)
            {
                var comparison = CompareValues(args[startIndex + i], lookupValue);
                if (comparison == 0)
                {
                    return i; // Exact match
                }
                else if (comparison < 0)
                {
                    lastMatch = i; // This is smaller, keep as candidate
                }
                else
                {
                    break; // We've passed the lookup value
                }
            }
        }
        else // Backward search
        {
            for (var i = length - 1; i >= 0; i--)
            {
                var comparison = CompareValues(args[startIndex + i], lookupValue);
                if (comparison == 0)
                {
                    return i; // Exact match
                }
                else if (comparison < 0)
                {
                    lastMatch = i; // This is smaller, keep as candidate
                }
                else
                {
                    break; // We've passed the lookup value
                }
            }
        }

        return lastMatch;
    }

    private static int FindExactOrNextLarger(CellValue[] args, int startIndex, int length, CellValue lookupValue, int searchMode)
    {
        if (searchMode == 1 || searchMode == 2) // Forward search
        {
            for (var i = 0; i < length; i++)
            {
                var comparison = CompareValues(args[startIndex + i], lookupValue);
                if (comparison == 0)
                {
                    return i; // Exact match
                }
                else if (comparison > 0)
                {
                    return i; // This is larger
                }
            }
        }
        else // Backward search
        {
            for (var i = length - 1; i >= 0; i--)
            {
                var comparison = CompareValues(args[startIndex + i], lookupValue);
                if (comparison == 0)
                {
                    return i; // Exact match
                }
                else if (comparison > 0)
                {
                    return i; // This is larger
                }
            }
        }

        return -1;
    }

    private static int FindWildcardMatch(CellValue[] args, int startIndex, int length, CellValue lookupValue, int searchMode)
    {
        if (lookupValue.Type != CellValueType.Text)
        {
            return -1; // Wildcard matching only works with text
        }

        var pattern = ConvertWildcardToRegex(lookupValue.StringValue);

        if (searchMode == 1) // First to last
        {
            for (var i = 0; i < length; i++)
            {
                var arrayValue = args[startIndex + i];
                if (arrayValue.Type == CellValueType.Text && System.Text.RegularExpressions.Regex.IsMatch(arrayValue.StringValue, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    return i;
                }
            }
        }
        else // Last to first
        {
            for (var i = length - 1; i >= 0; i--)
            {
                var arrayValue = args[startIndex + i];
                if (arrayValue.Type == CellValueType.Text && System.Text.RegularExpressions.Regex.IsMatch(arrayValue.StringValue, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                {
                    return i;
                }
            }
        }

        return -1;
    }

    private static string ConvertWildcardToRegex(string wildcardPattern)
    {
        // Excel wildcards: ? (single char), * (any chars), ~ (escape)
        var escaped = System.Text.RegularExpressions.Regex.Escape(wildcardPattern);
        escaped = escaped.Replace(@"\*", ".*");
        escaped = escaped.Replace(@"\?", ".");
        escaped = escaped.Replace(@"\~\*", @"\*");
        escaped = escaped.Replace(@"\~\?", @"\?");
        return "^" + escaped + "$";
    }

    private static int BinarySearch(CellValue[] args, int startIndex, int length, CellValue lookupValue, bool ascending)
    {
        var left = 0;
        var right = length - 1;

        while (left <= right)
        {
            var mid = left + (right - left) / 2;
            var comparison = CompareValues(args[startIndex + mid], lookupValue);

            if (!ascending)
            {
                comparison = -comparison; // Reverse comparison for descending order
            }

            if (comparison == 0)
            {
                return mid;
            }
            else if (comparison < 0)
            {
                left = mid + 1;
            }
            else
            {
                right = mid - 1;
            }
        }

        return -1;
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
