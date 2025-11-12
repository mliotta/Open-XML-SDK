// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the LCM function.
/// LCM(number1, [number2], ...) - returns the least common multiple of two or more integers.
/// </summary>
public sealed class LcmFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly LcmFunction Instance = new();

    private LcmFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "LCM";

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

                // LCM is only defined for positive integers
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

        // Calculate LCM of all numbers
        var result = numbers[0];
        for (int i = 1; i < numbers.Count; i++)
        {
            result = CalculateLcm(result, numbers[i]);

            // Check for overflow
            if (result < 0)
            {
                return CellValue.Error("#NUM!");
            }
        }

        return CellValue.FromNumber(result);
    }

    private static long CalculateLcm(long a, long b)
    {
        if (a == 0 || b == 0)
        {
            return 0;
        }

        // LCM(a,b) = abs(a*b) / GCD(a,b)
        var gcd = CalculateGcd(a, b);
        return System.Math.Abs((a / gcd) * b);
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
