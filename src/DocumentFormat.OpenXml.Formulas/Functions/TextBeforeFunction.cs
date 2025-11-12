// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the TEXTBEFORE function.
/// TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]) - extracts text before delimiter.
/// </summary>
public sealed class TextBeforeFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly TextBeforeFunction Instance = new();

    private TextBeforeFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "TEXTBEFORE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 6)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var text = args[0].StringValue;
        var delimiter = args[1].StringValue;

        // Default values
        var instanceNum = 1;
        var matchMode = 0; // 0 = case-sensitive, 1 = case-insensitive
        var matchEnd = 0; // 0 = search from start, 1 = search from end
        var ifNotFound = CellValue.Error("#N/A");

        // Parse optional arguments
        if (args.Length >= 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type == CellValueType.Number)
            {
                instanceNum = (int)args[2].NumericValue;
                if (instanceNum == 0 || instanceNum < -1)
                {
                    return CellValue.Error("#VALUE!");
                }
            }
        }

        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type == CellValueType.Number)
            {
                matchMode = (int)args[3].NumericValue;
                if (matchMode != 0 && matchMode != 1)
                {
                    return CellValue.Error("#VALUE!");
                }
            }
        }

        if (args.Length >= 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type == CellValueType.Number)
            {
                matchEnd = (int)args[4].NumericValue;
                if (matchEnd != 0 && matchEnd != 1)
                {
                    return CellValue.Error("#VALUE!");
                }
            }
        }

        if (args.Length >= 6)
        {
            if (!args[5].IsError)
            {
                ifNotFound = args[5];
            }
        }

        // Empty delimiter returns empty string
        if (string.IsNullOrEmpty(delimiter))
        {
            return CellValue.FromString(string.Empty);
        }

        var comparisonType = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

        // Search from end (match_end = 1)
        if (matchEnd == 1)
        {
            var absInstanceNum = instanceNum < 0 ? -instanceNum : instanceNum;
            var position = FindNthOccurrenceFromEnd(text, delimiter, absInstanceNum, comparisonType);
            if (position == -1)
            {
                return ifNotFound;
            }

            return CellValue.FromString(text.Substring(0, position));
        }

        // Search from start (match_end = 0)
        if (instanceNum > 0)
        {
            var position = FindNthOccurrence(text, delimiter, instanceNum, comparisonType);
            if (position == -1)
            {
                return ifNotFound;
            }

            return CellValue.FromString(text.Substring(0, position));
        }
        else
        {
            // Negative instance_num searches from end
            var absInstanceNum = instanceNum < 0 ? -instanceNum : instanceNum;
            var position = FindNthOccurrenceFromEnd(text, delimiter, absInstanceNum, comparisonType);
            if (position == -1)
            {
                return ifNotFound;
            }

            return CellValue.FromString(text.Substring(0, position));
        }
    }

    private static int FindNthOccurrence(string text, string delimiter, int n, StringComparison comparison)
    {
        var currentIndex = 0;
        for (var i = 0; i < n; i++)
        {
            var position = text.IndexOf(delimiter, currentIndex, comparison);
            if (position == -1)
            {
                return -1;
            }

            if (i == n - 1)
            {
                return position;
            }

            currentIndex = position + delimiter.Length;
        }

        return -1;
    }

    private static int FindNthOccurrenceFromEnd(string text, string delimiter, int n, StringComparison comparison)
    {
        var currentIndex = text.Length;
        for (var i = 0; i < n; i++)
        {
            var position = text.LastIndexOf(delimiter, currentIndex - 1, comparison);
            if (position == -1)
            {
                return -1;
            }

            if (i == n - 1)
            {
                return position;
            }

            currentIndex = position;
        }

        return -1;
    }
}
