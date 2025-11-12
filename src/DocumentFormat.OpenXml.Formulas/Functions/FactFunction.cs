// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FACT function.
/// FACT(number) - returns the factorial of a number.
/// </summary>
public sealed class FactFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FactFunction Instance = new();

    private FactFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FACT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
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

        // Factorial is only defined for non-negative integers
        if (number < 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Truncate to integer
        var n = (int)System.Math.Floor(number);

        // 0! = 1
        if (n == 0)
        {
            return CellValue.FromNumber(1);
        }

        // Calculate factorial iteratively
        double result = 1.0;
        for (int i = 2; i <= n; i++)
        {
            result *= i;

            // Check for overflow
            if (double.IsInfinity(result))
            {
                return CellValue.Error("#NUM!");
            }
        }

        return CellValue.FromNumber(result);
    }
}
