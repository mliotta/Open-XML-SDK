// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the SEARCH function.
/// SEARCH(find_text, within_text, [start_num]) - finds text (case-insensitive, wildcards, 1-based).
/// </summary>
public sealed class SearchFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly SearchFunction Instance = new();

    private SearchFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "SEARCH";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2 || args.Length > 3)
        {
            return FormulaResult.Error("#VALUE!");
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

            if (args[2].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            startNum = (int)args[2].NumericValue;

            if (startNum < 1)
            {
                return FormulaResult.Error("#VALUE!");
            }
        }

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        if (startIndex >= withinText.Length)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Convert Excel wildcards to regex (? = single char, * = any chars)
        var pattern = Regex.Escape(findText)
            .Replace(@"\?", ".")
            .Replace(@"\*", ".*");

        var match = Regex.Match(withinText.Substring(startIndex), pattern, RegexOptions.IgnoreCase);

        if (!match.Success)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Return 1-based position
        return FormulaResult.FromNumber(startIndex + match.Index + 1);
    }
}
