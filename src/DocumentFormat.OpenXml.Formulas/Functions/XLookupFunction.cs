// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the XLOOKUP function.
/// XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).
/// Modern replacement for VLOOKUP/HLOOKUP with more features.
/// match_mode: 0 (exact match, default), -1 (exact or next smaller), 1 (exact or next larger), 2 (wildcard).
/// search_mode: 1 (search first to last, default), -1 (search last to first), 2 (binary search ascending), -2 (binary search descending).
/// </summary>
public sealed class XLookupFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly XLookupFunction Instance = new();

    private XLookupFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "XLOOKUP";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3)
        {
            return CellValue.Error("#VALUE!");
        }

        // Parse arguments
        var lookupValue = args[0];
        if (lookupValue.IsError)
        {
            return lookupValue;
        }

        // Determine argument structure - need to find where lookup_array ends and return_array begins
        // Strategy: Look from the end backwards for optional parameters
        var argsLength = args.Length;
        var hasSearchMode = argsLength >= 6 && args[argsLength - 1].Type == CellValueType.Number;
        var hasMatchMode = argsLength >= 5 && args[argsLength - (hasSearchMode ? 2 : 1)].Type == CellValueType.Number;
        var hasIfNotFound = argsLength >= 4;

        // Calculate where arrays start/end
        var optionalParamsCount = (hasSearchMode ? 1 : 0) + (hasMatchMode ? 1 : 0) + (hasIfNotFound ? 1 : 0);
        var lookupArrayStart = 1;

        // For simplicity, assume equal-sized arrays split remaining args in half
        var remainingAfterLookupValue = argsLength - 1 - optionalParamsCount;
        if (remainingAfterLookupValue < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        var arrayLength = remainingAfterLookupValue / 2;
        var lookupArrayEnd = lookupArrayStart + arrayLength;
        var returnArrayStart = lookupArrayEnd;
        var returnArrayEnd = returnArrayStart + arrayLength;

        // Extract optional parameters
        CellValue ifNotFound = CellValue.Error("#N/A");
        var matchMode = 0;
        var searchMode = 1;

        var currentOptionalIndex = returnArrayEnd;
        if (hasIfNotFound && currentOptionalIndex < argsLength)
        {
            ifNotFound = args[currentOptionalIndex];
            currentOptionalIndex++;
        }

        if (hasMatchMode && currentOptionalIndex < argsLength)
        {
            var matchModeArg = args[currentOptionalIndex];
            if (matchModeArg.IsError)
            {
                return matchModeArg;
            }

            if (matchModeArg.Type == CellValueType.Number)
            {
                matchMode = (int)matchModeArg.NumericValue;
                if (matchMode < -1 || matchMode > 2)
                {
                    return CellValue.Error("#VALUE!");
                }
            }

            currentOptionalIndex++;
        }

        if (hasSearchMode && currentOptionalIndex < argsLength)
        {
            var searchModeArg = args[currentOptionalIndex];
            if (searchModeArg.IsError)
            {
                return searchModeArg;
            }

            if (searchModeArg.Type == CellValueType.Number)
            {
                searchMode = (int)searchModeArg.NumericValue;
                if (searchMode < -2 || searchMode > 2 || searchMode == 0)
                {
                    return CellValue.Error("#VALUE!");
                }
            }
        }

        // Check for errors in arrays
        for (var i = lookupArrayStart; i < returnArrayEnd; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Perform lookup based on match_mode and search_mode
        var matchIndex = -1;

        if (matchMode == 0) // Exact match
        {
            matchIndex = FindExactMatch(args, lookupArrayStart, arrayLength, lookupValue, searchMode);
        }
        else if (matchMode == -1) // Exact match or next smaller
        {
            matchIndex = FindExactOrNextSmaller(args, lookupArrayStart, arrayLength, lookupValue, searchMode);
        }
        else if (matchMode == 1) // Exact match or next larger
        {
            matchIndex = FindExactOrNextLarger(args, lookupArrayStart, arrayLength, lookupValue, searchMode);
        }
        else if (matchMode == 2) // Wildcard match
        {
            matchIndex = FindWildcardMatch(args, lookupArrayStart, arrayLength, lookupValue, searchMode);
        }

        if (matchIndex >= 0)
        {
            // Return corresponding value from return_array
            return args[returnArrayStart + matchIndex];
        }

        // No match found - return if_not_found value
        return ifNotFound;
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
