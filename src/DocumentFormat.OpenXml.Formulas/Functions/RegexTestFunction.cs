// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the REGEXTEST function.
/// REGEXTEST(text, pattern, [mode]) - Tests if text matches a regex pattern.
/// Returns TRUE if match found, FALSE otherwise.
/// Mode bitmask: 1=case-insensitive, 2=multiline, 4=singleline.
/// </summary>
public sealed class RegexTestFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RegexTestFunction Instance = new();

    private RegexTestFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "REGEXTEST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 2 || args.Length > 3)
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

        if (args.Length >= 3)
        {
            if (args[2].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            mode = (int)args[2].NumericValue;
            if (mode < 0)
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
            return FormulaResult.FromBool(Regex.IsMatch(text, pattern, options));
        }
        catch (ArgumentException)
        {
            return FormulaResult.Error("#VALUE!");
        }
    }
}
