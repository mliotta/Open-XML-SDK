// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the MULTINOMIAL function.
/// MULTINOMIAL(number1, [number2], ...) - returns the multinomial coefficient.
/// Formula: (n1 + n2 + ... + nk)! / (n1! * n2! * ... * nk!)
/// </summary>
public sealed class MultinomialFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly MultinomialFunction Instance = new();

    private MultinomialFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "MULTINOMIAL";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var numbers = new System.Collections.Generic.List<int>();
        var sum = 0;

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            var value = arg.NumericValue;

            // Must be non-negative
            if (value < 0)
            {
                return CellValue.Error("#NUM!");
            }

            // Truncate to integer
            var n = (int)System.Math.Floor(value);
            numbers.Add(n);
            sum += n;
        }

        // Calculate multinomial coefficient iteratively
        // (n1 + n2 + ... + nk)! / (n1! * n2! * ... * nk!)
        // Using iterative approach: result = C(sum, n1) * C(sum-n1, n2) * ... * C(nk, nk)
        double result = 1.0;
        var remaining = sum;

        foreach (var n in numbers)
        {
            // Calculate C(remaining, n) iteratively
            if (n > 0 && n < remaining)
            {
                var k = System.Math.Min(n, remaining - n);
                for (int i = 0; i < k; i++)
                {
                    result *= (remaining - i);
                    result /= (i + 1);

                    // Check for overflow
                    if (double.IsInfinity(result))
                    {
                        return CellValue.Error("#NUM!");
                    }
                }
            }
            else if (n == remaining)
            {
                // C(n, n) = 1, no change to result
            }

            remaining -= n;
        }

        return CellValue.FromNumber(System.Math.Round(result, 0));
    }
}
