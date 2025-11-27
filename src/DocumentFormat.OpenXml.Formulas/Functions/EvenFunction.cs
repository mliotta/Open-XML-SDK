// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the EVEN function.
/// EVEN(number) - rounds a number up to the nearest even integer.
/// </summary>
public sealed class EvenFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly EvenFunction Instance = new();

    private EvenFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "EVEN";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 1)
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

        // For positive numbers, round up to next even
        // For negative numbers, round down (away from zero) to next even
        double result;
        if (number >= 0)
        {
            result = System.Math.Ceiling(number);
            if (result % 2 != 0)
            {
                result += 1;
            }
        }
        else
        {
            result = System.Math.Floor(number);
            if (result % 2 != 0)
            {
                result -= 1;
            }
        }

        return FormulaResult.FromNumber(result);
    }
}
