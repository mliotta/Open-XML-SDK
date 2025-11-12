// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the XIRR function.
/// XIRR(values, dates, [guess]) - calculates the internal rate of return for a schedule of cash flows that is not necessarily periodic.
/// Uses Newton-Raphson method to solve for the rate where XNPV = 0.
/// </summary>
public sealed class XirrFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly XirrFunction Instance = new();

    private const double DefaultGuess = 0.1;
    private const double Tolerance = 1e-7;
    private const int MaxIterations = 100;

    private XirrFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "XIRR";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        // XIRR can be called with:
        // 1. Three arguments: values_range, dates_range, guess
        // 2. Multiple arguments where values and dates alternate, with optional guess at end
        // For this implementation, we'll support format: value1, date1, value2, date2, ..., [guess]

        if (args.Length < 4)
        {
            return CellValue.Error("#VALUE!");
        }

        var guess = DefaultGuess;
        var effectiveArgCount = args.Length;

        // Check if last argument might be the guess
        // If we have an odd number of args (excluding first if it's the rate), the last might be guess
        // For simplicity, we'll check if (args.Length - 1) is odd, last might be guess
        // Actually, let's keep it simple: last arg is guess if args.Length is odd
        if (args.Length % 2 != 0)
        {
            // Last argument is guess
            if (args[args.Length - 1].IsError)
            {
                return args[args.Length - 1];
            }

            if (args[args.Length - 1].Type == CellValueType.Number)
            {
                guess = args[args.Length - 1].NumericValue;
                effectiveArgCount = args.Length - 1;
            }
        }

        // Extract value-date pairs
        if (effectiveArgCount % 2 != 0)
        {
            return CellValue.Error("#VALUE!");
        }

        var pairCount = effectiveArgCount / 2;
        var values = new double[pairCount];
        var dates = new double[pairCount];

        for (int i = 0; i < pairCount; i++)
        {
            var valueIdx = i * 2;
            var dateIdx = i * 2 + 1;

            if (args[valueIdx].IsError)
            {
                return args[valueIdx];
            }

            if (args[dateIdx].IsError)
            {
                return args[dateIdx];
            }

            if (args[valueIdx].Type != CellValueType.Number || args[dateIdx].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            values[i] = args[valueIdx].NumericValue;
            dates[i] = args[dateIdx].NumericValue;
        }

        if (pairCount < 2)
        {
            return CellValue.Error("#VALUE!");
        }

        // XIRR requires at least one positive and one negative cash flow
        bool hasPositive = false;
        bool hasNegative = false;

        foreach (var value in values)
        {
            if (value > 0)
            {
                hasPositive = true;
            }
            else if (value < 0)
            {
                hasNegative = true;
            }

            if (hasPositive && hasNegative)
            {
                break;
            }
        }

        if (!hasPositive || !hasNegative)
        {
            return CellValue.Error("#NUM!");
        }

        var firstDate = dates[0];

        // Use Newton-Raphson method to find the rate where XNPV = 0
        var rate = guess;

        for (int iteration = 0; iteration < MaxIterations; iteration++)
        {
            // Calculate XNPV and its derivative at current rate
            double xnpv = 0.0;
            double dxnpv = 0.0; // derivative of XNPV with respect to rate

            for (int i = 0; i < pairCount; i++)
            {
                var yearFraction = (dates[i] - firstDate) / 365.0;
                var discountFactor = System.Math.Pow(1 + rate, yearFraction);

                if (double.IsInfinity(discountFactor) || double.IsNaN(discountFactor) || discountFactor == 0)
                {
                    return CellValue.Error("#NUM!");
                }

                // XNPV += value / (1 + rate)^yearFraction
                xnpv += values[i] / discountFactor;

                // Derivative: d/dr[value / (1+r)^yf] = -value * yf * (1+r)^(-yf-1)
                dxnpv -= values[i] * yearFraction / (discountFactor * (1 + rate));
            }

            // Check for convergence
            if (System.Math.Abs(xnpv) < Tolerance)
            {
                if (double.IsNaN(rate) || double.IsInfinity(rate))
                {
                    return CellValue.Error("#NUM!");
                }

                return CellValue.FromNumber(rate);
            }

            // Newton-Raphson iteration: rate_new = rate_old - f(rate) / f'(rate)
            if (System.Math.Abs(dxnpv) < 1e-10)
            {
                // Derivative too small, can't continue
                return CellValue.Error("#NUM!");
            }

            var newRate = rate - xnpv / dxnpv;

            // Prevent wild oscillations
            if (System.Math.Abs(newRate - rate) < Tolerance)
            {
                rate = newRate;
                break;
            }

            rate = newRate;

            // Bound the rate to prevent divergence
            if (rate < -0.99999)
            {
                rate = -0.99999;
            }
            else if (rate > 10.0)
            {
                rate = 10.0;
            }
        }

        // Final verification
        double finalXnpv = 0.0;
        for (int i = 0; i < pairCount; i++)
        {
            var yearFraction = (dates[i] - firstDate) / 365.0;
            var discountFactor = System.Math.Pow(1 + rate, yearFraction);
            finalXnpv += values[i] / discountFactor;
        }

        if (System.Math.Abs(finalXnpv) > 0.01)
        {
            // Solution didn't converge well enough
            return CellValue.Error("#NUM!");
        }

        if (double.IsNaN(rate) || double.IsInfinity(rate))
        {
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(rate);
    }
}
