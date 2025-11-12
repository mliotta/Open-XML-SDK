// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the COMBINA function.
/// COMBINA(number, number_chosen) - combinations with repetitions = C(n+k-1, k).
/// Formula: (n+k-1)! / (k! * (n-1)!)
/// </summary>
public sealed class CombinaFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly CombinaFunction Instance = new();

    private CombinaFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "COMBINA";

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

        // Special case: if n is 0, result is 1 only if k is 0
        if (n == 0)
        {
            return CellValue.FromNumber(k == 0 ? 1 : 0);
        }

        // COMBINA(n, k) = COMBIN(n + k - 1, k)
        // This is the number of ways to choose k items from n types with repetition
        var combinParam = n + k - 1;

        // Use symmetry property to minimize calculations
        if (k > combinParam - k)
        {
            k = combinParam - k;
        }

        // C(n+k-1, k) = 0 if k > n+k-1 (should not happen with our formula)
        if (k > combinParam)
        {
            return CellValue.Error("#NUM!");
        }

        // C(n+k-1, 0) = 1
        if (k == 0)
        {
            return CellValue.FromNumber(1);
        }

        // Calculate iteratively to avoid large factorials
        // C(combinParam, k) = combinParam * (combinParam-1) * ... * (combinParam-k+1) / (k * (k-1) * ... * 1)
        double result = 1.0;
        for (int i = 1; i <= k; i++)
        {
            result *= (combinParam - k + i);
            result /= i;

            // Check for overflow
            if (double.IsInfinity(result))
            {
                return CellValue.Error("#NUM!");
            }
        }

        return CellValue.FromNumber(System.Math.Round(result, 0));
    }
}
