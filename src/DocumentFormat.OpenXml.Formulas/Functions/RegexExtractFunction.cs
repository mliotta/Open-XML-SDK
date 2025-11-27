// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the REGEXEXTRACT function.
/// REGEXEXTRACT(text, pattern, [mode], [group]) - Extracts text matching a regex pattern.
/// Returns first match or specified capture group. Returns #N/A if no match.
/// </summary>
public sealed class RegexExtractFunction : IFunctionImplementation
{
    public static readonly RegexExtractFunction Instance = new();

    private RegexExtractFunction()
    {
    }

    public string Name => "REGEXEXTRACT";

    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2 || args.Length > 4)
        {
            return FormulaResult.Error("#VALUE!");
        }

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        if (args[0].Type != FormulaResultType.Text || args[1].Type != FormulaResultType.Text)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var text = args[0].StringValue;
        var pattern = args[1].StringValue;
        var mode = 0;
        var group = 0;

        if (args.Length >= 3 && args[2].Type == FormulaResultType.Number)
        {
            mode = (int)args[2].NumericValue;
        }

        if (args.Length >= 4 && args[3].Type == FormulaResultType.Number)
        {
            group = (int)args[3].NumericValue;
            if (group < 0)
            {
                return FormulaResult.Error("#VALUE!");
            }
        }

        var options = RegexOptions.None;
        if ((mode & 1) != 0) options |= RegexOptions.IgnoreCase;
        if ((mode & 2) != 0) options |= RegexOptions.Multiline;
        if ((mode & 4) != 0) options |= RegexOptions.Singleline;

        try
        {
            // Use static method which caches compiled regexes internally
            var match = Regex.Match(text, pattern, options);

            if (!match.Success)
            {
                return FormulaResult.Error("#N/A");
            }

            if (group >= match.Groups.Count)
            {
                return FormulaResult.Error("#VALUE!");
            }

            return FormulaResult.FromString(match.Groups[group].Value);
        }
        catch (ArgumentException)
        {
            return FormulaResult.Error("#VALUE!");
        }
    }
}
