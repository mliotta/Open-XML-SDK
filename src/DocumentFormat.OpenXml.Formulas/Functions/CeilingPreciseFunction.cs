// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CEILING.PRECISE function.
/// CEILING.PRECISE(number, [significance]) - always rounds up regardless of sign.
/// </summary>
public sealed class CeilingPreciseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CeilingPreciseFunction Instance = new();

    private CeilingPreciseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CEILING.PRECISE";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }

        var number = args[0].NumericValue;

        // Default significance is 1 if not provided
        double significance = 1;
        if (args.Length == 2)
        {
            if (args[1].IsError)
            {
                return args[1];
            }

            if (args[1].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            significance = args[1].NumericValue;
        }

        if (significance == 0)
        {
            return FormulaResult.FromNumber(0);
        }

        // Use absolute value of significance
        significance = System.Math.Abs(significance);

        // CEILING.PRECISE always rounds toward positive infinity
        // For positive numbers: 4.3 -> 5 (up, away from zero)
        // For negative numbers: -4.3 -> -4 (up, toward zero)
        double result = System.Math.Ceiling(number / significance) * significance;

        return FormulaResult.FromNumber(result);
    }
}
