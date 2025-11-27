// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the IRR function.
/// IRR(values, [guess]) - calculates the internal rate of return for a series of cash flows.
/// </summary>
public sealed class IrrFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly IrrFunction Instance = new();

    private const double DefaultGuess = 0.1;
    private const double Tolerance = 1e-7;
    private const int MaxIterations = 100;

    private IrrFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "IRR";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        // IRR needs at least 2 cash flow values to calculate a return rate
        if (args.Length < 2)
        {
            return FormulaResult.Error("#VALUE!");
        }

        // Check for errors in all arguments first
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].IsError)
            {
                return args[i];
            }
        }

        // Extract cash flow values - must be an array or range
        double[] values;

        // For this implementation, we treat all arguments as cash flow values
        // A proper implementation would also support an optional guess parameter

        // Re-interpret: args are value1, value2, ..., [guess]
        var guess = DefaultGuess;

        // Extract all values
        values = new double[args.Length];

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].Type != FormulaResultType.Number)
            {
                return FormulaResult.Error("#VALUE!");
            }

            values[i] = args[i].NumericValue;
        }

        // IRR requires at least one positive and one negative cash flow
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
            return FormulaResult.Error("#NUM!");
        }

        // Use Newton-Raphson method to find the rate where NPV = 0
        var rate = guess;

        for (int iteration = 0; iteration < MaxIterations; iteration++)
        {
            // Calculate NPV and its derivative at current rate
            // IRR NPV formula: sum of values[i] / (1+r)^i for i = 0, 1, 2, ...
            double npv = 0.0;
            double dnpv = 0.0; // derivative of NPV with respect to rate

            for (int i = 0; i < values.Length; i++)
            {
                var period = i; // Period starts at 0 for IRR calculation
                var discountFactor = System.Math.Pow(1 + rate, period);

                if (double.IsInfinity(discountFactor) || double.IsNaN(discountFactor))
                {
                    return FormulaResult.Error("#NUM!");
                }

                // NPV += value / (1 + rate)^period
                npv += values[i] / discountFactor;

                // Derivative: d/dr[value / (1+r)^p] = -value * p * (1+r)^(-p-1)
                if (period > 0)
                {
                    dnpv -= values[i] * period / (discountFactor * (1 + rate));
                }
            }

            // Check for convergence
            if (System.Math.Abs(npv) < Tolerance)
            {
                if (double.IsNaN(rate) || double.IsInfinity(rate))
                {
                    return FormulaResult.Error("#NUM!");
                }

                return FormulaResult.FromNumber(rate);
            }

            // Newton-Raphson iteration: rate_new = rate_old - f(rate) / f'(rate)
            if (System.Math.Abs(dnpv) < 1e-10)
            {
                // Derivative too small, can't continue
                return FormulaResult.Error("#NUM!");
            }

            var newRate = rate - npv / dnpv;

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
        double finalNpv = 0.0;
        for (int i = 0; i < values.Length; i++)
        {
            var period = i; // Period starts at 0 for IRR calculation
            var discountFactor = System.Math.Pow(1 + rate, period);
            finalNpv += values[i] / discountFactor;
        }

        if (System.Math.Abs(finalNpv) > 0.01)
        {
            // Solution didn't converge well enough
            return FormulaResult.Error("#NUM!");
        }

        if (double.IsNaN(rate) || double.IsInfinity(rate))
        {
            return FormulaResult.Error("#NUM!");
        }

        return FormulaResult.FromNumber(rate);
    }
}
