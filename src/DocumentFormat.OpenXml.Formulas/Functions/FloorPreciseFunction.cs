// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FLOOR.PRECISE function.
/// FLOOR.PRECISE(number, [significance]) - always rounds down regardless of sign.
/// </summary>
public sealed class FloorPreciseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FloorPreciseFunction Instance = new();

    private FloorPreciseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FLOOR.PRECISE";

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

        // Use absolute value of significance
        significance = System.Math.Abs(significance);

        // Always round toward negative infinity regardless of sign
        double result = System.Math.Floor(number / significance) * significance;

        return CellValue.FromNumber(result);
    }
}
