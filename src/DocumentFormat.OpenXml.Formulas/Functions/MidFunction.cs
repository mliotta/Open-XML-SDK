// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MID function.
/// MID(text, start_num, num_chars) - returns substring (1-based indexing).
/// </summary>
public sealed class MidFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MidFunction Instance = new();

    private MidFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MID";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 3)
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

        if (args[2].IsError)
        {
            return args[2];
        }

        var text = args[0].StringValue;

        if (args[1].Type != FormulaResultType.Number || args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var startNum = (int)args[1].NumericValue;
        var numChars = (int)args[2].NumericValue;

        if (startNum < 1 || numChars < 0)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Excel uses 1-based indexing
        var startIndex = startNum - 1;

        if (startIndex >= text.Length)
        {
            return FormulaResult.FromString(string.Empty);
        }

        var length = System.Math.Min(numChars, text.Length - startIndex);
        var result = text.Substring(startIndex, length);
        return FormulaResult.FromString(result);
    }
}
