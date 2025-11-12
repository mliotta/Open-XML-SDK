// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ISO.CEILING function.
/// ISO.CEILING(number, [significance]) - ISO standard ceiling (always rounds toward positive infinity).
/// </summary>
public sealed class IsoCeilingFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IsoCeilingFunction Instance = new();

    private IsoCeilingFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ISO.CEILING";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
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

            if (args[1].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            significance = args[1].NumericValue;
        }

        if (significance == 0)
        {
            return CellValue.FromNumber(0);
        }

        // Use absolute value of significance (ISO standard)
        significance = System.Math.Abs(significance);

        // Always round toward positive infinity regardless of sign
        double result;
        if (number >= 0)
        {
            result = System.Math.Ceiling(number / significance) * significance;
        }
        else
        {
            // For negative numbers, ceiling toward positive infinity means toward zero
            result = System.Math.Ceiling(System.Math.Abs(number) / significance) * significance * -1;
        }

        return CellValue.FromNumber(result);
    }
}
