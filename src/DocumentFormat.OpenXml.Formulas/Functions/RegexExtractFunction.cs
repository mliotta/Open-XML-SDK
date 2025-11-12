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

    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 2 || args.Length > 4)
        {
            return CellValue.Error("#VALUE!");
        }

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        if (args[0].Type != CellValueType.Text || args[1].Type != CellValueType.Text)
        {
            return CellValue.Error("#VALUE!");
        }

        var text = args[0].StringValue;
        var pattern = args[1].StringValue;
        var mode = 0;
        var group = 0;

        if (args.Length >= 3 && args[2].Type == CellValueType.Number)
        {
            mode = (int)args[2].NumericValue;
        }

        if (args.Length >= 4 && args[3].Type == CellValueType.Number)
        {
            group = (int)args[3].NumericValue;
            if (group < 0)
            {
                return CellValue.Error("#VALUE!");
            }
        }

        var options = RegexOptions.None;
        if ((mode & 1) != 0) options |= RegexOptions.IgnoreCase;
        if ((mode & 2) != 0) options |= RegexOptions.Multiline;
        if ((mode & 4) != 0) options |= RegexOptions.Singleline;

        try
        {
            var regex = new Regex(pattern, options);
            var match = regex.Match(text);

            if (!match.Success)
            {
                return CellValue.Error("#N/A");
            }

            if (group >= match.Groups.Count)
            {
                return CellValue.Error("#VALUE!");
            }

            return CellValue.FromString(match.Groups[group].Value);
        }
        catch (ArgumentException)
        {
            return CellValue.Error("#VALUE!");
        }
    }
}
