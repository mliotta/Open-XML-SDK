// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FLOOR.MATH function.
/// FLOOR.MATH(number, [significance], [mode]) - rounds down to the nearest multiple with mode parameter.
/// Mode: 0 (default) = round negative away from 0, 1 = toward 0
/// </summary>
public sealed class FloorMathFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FloorMathFunction Instance = new();

    private FloorMathFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FLOOR.MATH";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 3)
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
        if (args.Length >= 2)
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

        // Default mode is 0
        double mode = 0;
        if (args.Length == 3)
        {
            if (args[2].IsError)
            {
                return args[2];
            }

            if (args[2].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            mode = args[2].NumericValue;
        }

        if (significance == 0)
        {
            return CellValue.FromNumber(0);
        }

        // Use absolute value of significance
        significance = System.Math.Abs(significance);

        double result;
        if (number >= 0)
        {
            // Positive numbers always round down
            result = System.Math.Floor(number / significance) * significance;
        }
        else
        {
            // Negative numbers: mode determines direction
            if (mode == 0)
            {
                // Mode 0: round away from zero (down for negative numbers)
                result = System.Math.Floor(number / significance) * significance;
            }
            else
            {
                // Mode 1: round toward zero (up for negative numbers)
                result = System.Math.Floor(System.Math.Abs(number) / significance) * significance * -1;
            }
        }

        return CellValue.FromNumber(result);
    }
}
