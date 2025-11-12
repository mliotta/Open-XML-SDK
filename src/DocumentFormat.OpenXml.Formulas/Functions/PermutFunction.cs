// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the PERMUT function.
/// PERMUT(number, number_chosen) - returns the number of permutations for a given number of items.
/// Formula: n! / (n-k)! where k ≤ n
/// </summary>
public sealed class PermutFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly PermutFunction Instance = new();

    private PermutFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "PERMUT";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 2)
        {
            return CellValue.Error("#VALUE!");
        }

        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var number = args[0].NumericValue;
        var numberChosen = args[1].NumericValue;

        // Both arguments must be non-negative
        if (number < 0 || numberChosen < 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Truncate to integers
        var n = (int)System.Math.Floor(number);
        var k = (int)System.Math.Floor(numberChosen);

        // k must be <= n
        if (k > n)
        {
            return CellValue.Error("#NUM!");
        }

        // P(n, 0) = 1
        if (k == 0)
        {
            return CellValue.FromNumber(1);
        }

        // Calculate permutation iteratively
        // P(n, k) = n * (n-1) * ... * (n-k+1) = n! / (n-k)!
        double result = 1.0;
        for (int i = 0; i < k; i++)
        {
            result *= (n - i);

            // Check for overflow
            if (double.IsInfinity(result))
            {
                return CellValue.Error("#NUM!");
            }
        }

        return CellValue.FromNumber(result);
    }
}
