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
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Extract lookup_value (first argument)
        var lookupValue = args[0];

        // Validate lookup value
        if (lookupValue.IsError)
        {
            return lookupValue;
        }

        // XMATCH arguments:
        // args[0] = lookup_value
        // args[1..n] = lookup_array (variable length)
        // args[n+1] = match_mode (optional)
        // args[n+2] = search_mode (optional)

        // Strategy: Work backwards to identify optional parameters
        // Optional parameters are always at the end in this order: [match_mode], [search_mode]
        var matchMode = 0;
        var searchMode = 1;
        var hasMatchMode = false;
        var hasSearchMode = false;

        // Minimum: 1 lookup_value + 1 array element = 2 args
        // With all optionals: lookup_value + array + match_mode + search_mode = at least 4 args

        // Step 1: Check if we have both match_mode and search_mode (4+ args, last two are numbers)
        if (args.Length >= 4)
        {
            var lastArg = args[args.Length - 1];
            var secondLastArg = args[args.Length - 2];

            if (lastArg.IsError)
            {
                return lastArg;
            }

            if (secondLastArg.IsError)
            {
                return secondLastArg;
            }

            // Check if both last args are numbers
            if (lastArg.Type == FormulaResultType.Number && secondLastArg.Type == FormulaResultType.Number)
            {
                var lastVal = (int)lastArg.NumericValue;
                var secondLastVal = (int)secondLastArg.NumericValue;

                // Check if they form a valid (match_mode, search_mode) pair
                var isValidMatchMode = secondLastVal >= -1 && secondLastVal <= 2;
                var isValidSearchMode = lastVal >= -2 && lastVal <= 2 && lastVal != 0;

                if (isValidMatchMode && isValidSearchMode)
                {
                    // Both are valid - treat as match_mode + search_mode
                    hasMatchMode = true;
                    hasSearchMode = true;
                    matchMode = secondLastVal;
                    searchMode = lastVal;
                }
                else if (!isValidSearchMode && isValidMatchMode)
                {
                    // Last is invalid search_mode but second-last is valid match_mode
                    // This means user provided invalid search_mode
                    return FormulaResult.Error("#VALUE!");
                }
                // If isValidSearchMode but not isValidMatchMode, fall through to check just search_mode
                // If neither valid, fall through to check just match_mode
            }
        }

        // Step 2: If we didn't find both, check for just match_mode (3+ args, last is number)
        if (!hasMatchMode && !hasSearchMode && args.Length >= 3)
        {
            var lastArg = args[args.Length - 1];

            if (lastArg.IsError)
            {
                return lastArg;
            }

            if (lastArg.Type == FormulaResultType.Number)
            {
                var val = (int)lastArg.NumericValue;

                // Check if it's a valid match_mode
                if (val >= -1 && val <= 2)
                {
                    hasMatchMode = true;
                    matchMode = val;
                }
                else
                {
                    // It's a number but not valid match_mode - error
                    return FormulaResult.Error("#VALUE!");
                }
            }
        }

        var optionalCount = (hasSearchMode ? 1 : 0) + (hasMatchMode ? 1 : 0);
        var arrayLength = args.Length - 1 - optionalCount;

        if (arrayLength <= 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var arrayStartIndex = 1;

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
            return FormulaResult.FromNumber(matchIndex + 1);
        }

        // No match found
        return FormulaResult.Error("#N/A");
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
