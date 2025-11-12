// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the FACTDOUBLE function.
/// FACTDOUBLE(number) - double factorial (n!! = n * (n-2) * (n-4) * ... * 2 or 1).
/// </summary>
public sealed class FactDoubleFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly FactDoubleFunction Instance = new();

    private FactDoubleFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "FACTDOUBLE";

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

        // Double factorial is only defined for non-negative integers
        if (number < 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Truncate to integer
        var n = (int)System.Math.Floor(number);

        // 0!! = 1 and 1!! = 1
        if (n == 0 || n == 1)
        {
            return CellValue.FromNumber(1);
        }

        // Calculate double factorial iteratively
        // For even n: n!! = n * (n-2) * (n-4) * ... * 4 * 2
        // For odd n: n!! = n * (n-2) * (n-4) * ... * 3 * 1
        double result = 1.0;
        for (int i = n; i > 0; i -= 2)
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
