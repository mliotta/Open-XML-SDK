// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GCD function.
/// GCD(number1, [number2], ...) - returns the greatest common divisor of two or more integers.
/// </summary>
public sealed class GcdFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GcdFunction Instance = new();

    private GcdFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GCD";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Collect all numbers
        var numbers = new System.Collections.Generic.List<long>();

        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg; // Propagate errors
            }

            if (arg.Type == CellValueType.Number)
            {
                var value = System.Math.Abs(arg.NumericValue);
                var intValue = (long)System.Math.Floor(value);

                // GCD is only defined for integers
                if (intValue < 0)
                {
                    return CellValue.Error("#NUM!");
                }

                numbers.Add(intValue);
            }
            else
            {
                return CellValue.Error("#VALUE!");
            }
        }

        if (numbers.Count == 0)
        {
            return CellValue.Error("#VALUE!");
        }

        // Calculate GCD of all numbers
        var result = numbers[0];
        for (int i = 1; i < numbers.Count; i++)
        {
            result = CalculateGcd(result, numbers[i]);
        }

        return CellValue.FromNumber(result);
    }

    private static long CalculateGcd(long a, long b)
    {
        // Euclidean algorithm
        while (b != 0)
        {
            var temp = b;
            b = a % b;
            a = temp;
        }

        return a;
    }
}
