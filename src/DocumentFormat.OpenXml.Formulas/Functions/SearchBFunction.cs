// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SEARCHB function.
/// SEARCHB(find_text, within_text, [start_num]) - finds text by byte position (case-insensitive, wildcards, UTF-8, 1-based).
/// </summary>
public sealed class SearchBFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SearchBFunction Instance = new();

    private SearchBFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SEARCHB";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 3)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        var findText = args[0].StringValue;
        var withinText = args[1].StringValue;
        var startNum = 1;

        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            startNum = (int)args[2].NumericValue;

            if (startNum < 1)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        // Convert to bytes for position calculation
        var withinBytes = Encoding.UTF8.GetBytes(withinText);

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        if (startIndex >= withinBytes.Length)
        {
            return CellValue.Error("#VALUE!");
        }

        // Convert start byte position to character position for regex matching
        var charStartIndex = Encoding.UTF8.GetString(withinBytes, 0, startIndex).Length;

        // Convert Excel wildcards to regex (? = single char, * = any chars)
        var pattern = Regex.Escape(findText)
            .Replace(@"\?", ".")
            .Replace(@"\*", ".*");

        var searchText = withinText.Substring(charStartIndex);
        var match = Regex.Match(searchText, pattern, RegexOptions.IgnoreCase);

        if (!match.Success)
        {
            return CellValue.Error("#VALUE!");
        }

        // Calculate the byte position of the match
        var matchCharPosition = charStartIndex + match.Index;
        var textUpToMatch = withinText.Substring(0, matchCharPosition);
        var bytePosition = Encoding.UTF8.GetByteCount(textUpToMatch);

        // Return 1-based position
        return CellValue.FromNumber(bytePosition + 1);
    }
}
