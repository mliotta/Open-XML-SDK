// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the REGEXREPLACE function.
/// REGEXREPLACE(text, pattern, replacement, [mode], [occurrence]) - Replaces text matching regex pattern.
/// </summary>
public sealed class RegexReplaceFunction : IFunctionImplementation
{
    public static readonly RegexReplaceFunction Instance = new();

    private RegexReplaceFunction()
    {
    }

    public string Name => "REGEXREPLACE";

    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 5)
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

        if (args[0].Type != CellValueType.Text || args[1].Type != CellValueType.Text || args[2].Type != CellValueType.Text)
        {
            return CellValue.Error("#VALUE!");
        }

        var text = args[0].StringValue;
        var pattern = args[1].StringValue;
        var replacement = args[2].StringValue;
        var mode = 0;
        var occurrence = 0;

        if (args.Length >= 4 && args[3].Type == CellValueType.Number)
        {
            mode = (int)args[3].NumericValue;
        }

        if (args.Length >= 5 && args[4].Type == CellValueType.Number)
        {
            occurrence = (int)args[4].NumericValue;
            if (occurrence < 0)
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
            string result;
            if (occurrence == 0)
            {
                // Use static method which caches compiled regexes internally
                result = Regex.Replace(text, pattern, replacement, options);
            }
            else
            {
                // For selective replacement, still need instance (can't use static method with evaluator)
                var regex = new Regex(pattern, options);
                var count = 0;
                result = regex.Replace(text, match =>
                {
                    count++;
                    return count == occurrence ? replacement : match.Value;
                });
            }

            return CellValue.FromString(result);
        }
        catch (ArgumentException)
        {
            return CellValue.Error("#VALUE!");
        }
    }
}
