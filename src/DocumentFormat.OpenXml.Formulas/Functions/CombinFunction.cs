// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COMBIN function.
/// COMBIN(number, number_chosen) - returns the number of combinations for a given number of items.
/// Formula: n! / (k! * (n-k)!) where k ≤ n
/// </summary>
public sealed class CombinFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CombinFunction Instance = new();

    private CombinFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COMBIN";

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
        var numberChosen = args[1].NumericValue;

        // Both arguments must be non-negative
        if (number < 0 || numberChosen < 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Truncate to integers
        var n = (int)System.Math.Floor(number);
        var k = (int)System.Math.Floor(numberChosen);

        // k must be <= n
        if (k > n)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Calculate combination using efficient method
        // C(n, k) = n! / (k! * (n-k)!)
        // Optimize by using C(n, k) = C(n, n-k) to minimize calculations
        if (k > n - k)
        {
            k = n - k;
        }

        // C(n, 0) = 1
        if (k == 0)
        {
            return FormulaResult.FromNumber(1);
        }

        // Calculate iteratively to avoid large factorials
        // C(n, k) = n * (n-1) * ... * (n-k+1) / (k * (k-1) * ... * 1)
        double result = 1.0;
        for (int i = 1; i <= k; i++)
        {
            result *= (n - k + i);
            result /= i;

            // Check for overflow
            if (double.IsInfinity(result))
            {
                return FormulaResult.Error("#NUM!");
            }
        }

        return FormulaResult.FromNumber(System.Math.Round(result, 0));
    }
}
