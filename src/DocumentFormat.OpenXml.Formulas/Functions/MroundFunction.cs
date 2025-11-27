// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MROUND function.
/// MROUND(number, multiple) - rounds a number to the nearest multiple.
/// Uses Excel's rounding mode (round half away from zero).
/// </summary>
public sealed class MroundFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MroundFunction Instance = new();

    private MroundFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MROUND";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 2)
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

        if (args[0].Type != FormulaResultType.Number || args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var number = args[0].NumericValue;
        var multiple = args[1].NumericValue;

        // Multiple cannot be zero
        if (multiple == 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Excel MROUND requires number and multiple to have the same sign
        if ((number > 0 && multiple < 0) || (number < 0 && multiple > 0))
        {
            return FormulaResult.Error("#NUM!");
        }

        // Calculate: ROUND(number/multiple, 0) * multiple
        // Use MidpointRounding.AwayFromZero to match Excel behavior
        var result = System.Math.Round(number / multiple, 0, MidpointRounding.AwayFromZero) * multiple;

        return FormulaResult.FromNumber(result);
    }
}
