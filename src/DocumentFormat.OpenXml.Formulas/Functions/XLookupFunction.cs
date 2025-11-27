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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 3)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Parse arguments
        var lookupValue = args[0];
        if (lookupValue.IsError)
        {
            return lookupValue;
        }

        // XLOOKUP has specific argument positions:
        // args[0] = lookup_value
        // args[1..n] = lookup_array (variable length)
        // args[n+1..m] = return_array (same length as lookup_array)
        // args[m+1] = if_not_found (optional)
        // args[m+2] = match_mode (optional)
        // args[m+3] = search_mode (optional)

        // Strategy: Try different combinations of optional parameters to find
        // a configuration where the remaining args can be split into two equal arrays

        var matchMode = 0;
        var searchMode = 1;
        FormulaResult ifNotFound = FormulaResult.Error("#N/A");
        var arrayLength = 0;
        var foundValidConfig = false;

        // Try configurations from most optional params to least
        // Config: (hasSearchMode, hasMatchMode, hasIfNotFound)
        var configs = new[]
        {
            new { hasSearchMode = true, hasMatchMode = true, hasIfNotFound = true },   // All 3 optional params
            new { hasSearchMode = false, hasMatchMode = true, hasIfNotFound = true },  // match_mode and if_not_found
            new { hasSearchMode = true, hasMatchMode = false, hasIfNotFound = true },  // search_mode and if_not_found (invalid - skip)
            new { hasSearchMode = false, hasMatchMode = false, hasIfNotFound = true }, // just if_not_found
            new { hasSearchMode = true, hasMatchMode = true, hasIfNotFound = false },  // match_mode and search_mode
            new { hasSearchMode = false, hasMatchMode = true, hasIfNotFound = false }, // just match_mode
            new { hasSearchMode = true, hasMatchMode = false, hasIfNotFound = false }, // just search_mode (invalid - skip)
            new { hasSearchMode = false, hasMatchMode = false, hasIfNotFound = false } // no optional params
        };

        foreach (var config in configs)
        {
            var hasSearchMode = config.hasSearchMode;
            var hasMatchMode = config.hasMatchMode;
            var hasIfNotFound = config.hasIfNotFound;

            // Search mode cannot exist without match mode
            if (hasSearchMode && !hasMatchMode)
            {
                continue;
            }

            var optionalCount = (hasSearchMode ? 1 : 0) + (hasMatchMode ? 1 : 0) + (hasIfNotFound ? 1 : 0);

            // Check if we have enough args for this configuration
            if (args.Length < 3 + optionalCount)
            {
                continue;
            }

            // Check if remaining args can be split into two equal arrays
            var totalArrayElements = args.Length - 1 - optionalCount;
            if (totalArrayElements % 2 != 0 || totalArrayElements < 2)
            {
                continue;
            }

            // Validate the optional parameters if they're supposed to be present
            var currentOptional = args.Length - optionalCount;
            var isValid = true;

            if (hasIfNotFound)
            {
                // if_not_found can be any type, always valid
                ifNotFound = args[currentOptional];
                currentOptional++;
            }

            if (hasMatchMode)
            {
                if (args[currentOptional].Type != FormulaResultType.Number)
                {
                    isValid = false;
                }
                else
                {
                    var val = (int)args[currentOptional].NumericValue;
                    if (val < -1 || val > 2)
                    {
                        isValid = false;
                    }
                    else
                    {
                        matchMode = val;
                    }
                }
                currentOptional++;
            }

            if (hasSearchMode && isValid)
            {
                if (args[currentOptional].Type != FormulaResultType.Number)
                {
                    isValid = false;
                }
                else
                {
                    var val = (int)args[currentOptional].NumericValue;
                    if (val < -2 || val > 2 || val == 0)
                    {
                        isValid = false;
                    }
                    else
                    {
                        searchMode = val;
                    }
                }
            }

            if (isValid)
            {
                arrayLength = totalArrayElements / 2;
                foundValidConfig = true;
                break;
            }
        }

        if (!foundValidConfig)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var lookupArrayStart = 1;
        var returnArrayStart = 1 + arrayLength;

        // Check for errors in arrays
        for (var i = lookupArrayStart; i < lookupArrayStart + arrayLength; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        for (var i = returnArrayStart; i < returnArrayStart + arrayLength; i++)
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

    private static int FindExactMatch(FormulaResult[] args, int startIndex, int length, FormulaResult lookupValue, int searchMode)
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

    private static int FindExactOrNextSmaller(FormulaResult[] args, int startIndex, int length, FormulaResult lookupValue, int searchMode)
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

    private static int FindExactOrNextLarger(FormulaResult[] args, int startIndex, int length, FormulaResult lookupValue, int searchMode)
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

    private static int FindWildcardMatch(FormulaResult[] args, int startIndex, int length, FormulaResult lookupValue, int searchMode)
    {
        if (lookupValue.Type != FormulaResultType.Text)
        {
            return -1; // Wildcard matching only works with text
        }

        var pattern = ConvertWildcardToRegex(lookupValue.StringValue);

        if (searchMode == 1) // First to last
        {
            for (var i = 0; i < length; i++)
            {
                var arrayValue = args[startIndex + i];
                if (arrayValue.Type == FormulaResultType.Text && System.Text.RegularExpressions.Regex.IsMatch(arrayValue.StringValue, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
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
                if (arrayValue.Type == FormulaResultType.Text && System.Text.RegularExpressions.Regex.IsMatch(arrayValue.StringValue, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
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

    private static int BinarySearch(FormulaResult[] args, int startIndex, int length, FormulaResult lookupValue, bool ascending)
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

    private static bool ValuesEqual(FormulaResult a, FormulaResult b)
    {
        if (a.Type != b.Type)
        {
            return false;
        }

        return a.Type switch
        {
            FormulaResultType.Number => System.Math.Abs(a.NumericValue - b.NumericValue) < 1e-10,
            FormulaResultType.Text => string.Equals(a.StringValue, b.StringValue, StringComparison.OrdinalIgnoreCase),
            FormulaResultType.Boolean => a.BoolValue == b.BoolValue,
            FormulaResultType.Empty => true,
            _ => false,
        };
    }

    private static int CompareValues(FormulaResult a, FormulaResult b)
    {
        // Compare two values for ordering
        if (a.Type != b.Type)
        {
            // Type mismatch - use type priority: Number < Text < Boolean < Empty
            return a.Type.CompareTo(b.Type);
        }

        return a.Type switch
        {
            FormulaResultType.Number => a.NumericValue.CompareTo(b.NumericValue),
            FormulaResultType.Text => string.Compare(a.StringValue, b.StringValue, StringComparison.OrdinalIgnoreCase),
            FormulaResultType.Boolean => a.BoolValue.CompareTo(b.BoolValue),
            FormulaResultType.Empty => 0,
            _ => 0,
        };
    }
}
