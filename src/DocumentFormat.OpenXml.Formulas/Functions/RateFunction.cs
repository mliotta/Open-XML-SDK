// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the RATE function.
/// RATE(nper, pmt, pv, [fv], [type], [guess]) - calculates the interest rate per period of an annuity.
/// </summary>
public sealed class RateFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly RateFunction Instance = new();

    private const double DefaultGuess = 0.1;
    private const double Tolerance = 1e-7;
    private const int MaxIterations = 100;

    private RateFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "RATE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 3 || args.Length > 6)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in required arguments
        if (args[0].IsError)
        {
            return args[0];
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[2].IsError)
        {
            return args[2];
        }

        // Validate required arguments are numbers
        if (args[0].Type != CellValueType.Number || args[1].Type != CellValueType.Number || args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var nper = args[0].NumericValue;
        var pmt = args[1].NumericValue;
        var pv = args[2].NumericValue;
        var fv = 0.0;
        var type = 0.0;
        var guess = DefaultGuess;

        // Optional fv parameter
        if (args.Length >= 4)
        {
            if (args[3].IsError)
            {
                return args[3];
            }

            if (args[3].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            fv = args[3].NumericValue;
        }

        // Optional type parameter
        if (args.Length >= 5)
        {
            if (args[4].IsError)
            {
                return args[4];
            }

            if (args[4].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            type = args[4].NumericValue;
        }

        // Optional guess parameter
        if (args.Length == 6)
        {
            if (args[5].IsError)
            {
                return args[5];
            }

            if (args[5].Type != CellValueType.Number)
            {
                return CellValue.Error("#VALUE!");
            }

            guess = args[5].NumericValue;
        }

        // Validate type is 0 or 1
        if (type != 0.0 && type != 1.0)
        {
            return CellValue.Error("#NUM!");
        }

        // Validate nper is positive
        if (nper <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Use Newton-Raphson method to find the rate
        var rate = guess;

        for (int i = 0; i < MaxIterations; i++)
        {
            // Calculate f(rate) - the financial equation that should equal zero
            double f;
            double df; // derivative of f

            if (System.Math.Abs(rate) < 1e-10)
            {
                // When rate is near zero, use simplified formulas
                f = pv + pmt * nper + fv;
                df = 0;

                // If f is close to zero, we found the solution
                if (System.Math.Abs(f) < Tolerance)
                {
                    return CellValue.FromNumber(0.0);
                }

                // Otherwise, start with a small non-zero rate
                rate = 0.01;
                continue;
            }

            var pow1 = System.Math.Pow(1 + rate, nper);
            var pow2 = System.Math.Pow(1 + rate, nper - 1);

            // f(rate) = pv * (1+rate)^nper + pmt * (1 + rate*type) * ((1+rate)^nper - 1) / rate + fv
            f = pv * pow1 + pmt * (1 + rate * type) * (pow1 - 1) / rate + fv;

            // Derivative df/drate
            df = pv * nper * pow2
                 + pmt * (1 + rate * type) * (nper * pow2 / rate - (pow1 - 1) / (rate * rate))
                 + pmt * type * (pow1 - 1) / rate;

            // Check for convergence
            if (System.Math.Abs(f) < Tolerance)
            {
                if (double.IsNaN(rate) || double.IsInfinity(rate))
                {
                    return CellValue.Error("#NUM!");
                }

                return CellValue.FromNumber(rate);
            }

            // Newton-Raphson iteration
            if (System.Math.Abs(df) < 1e-10)
            {
                // Derivative too small, can't continue
                return CellValue.Error("#NUM!");
            }

            var newRate = rate - f / df;

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

        // Check if we found a valid solution
        if (double.IsNaN(rate) || double.IsInfinity(rate))
        {
            return CellValue.Error("#NUM!");
        }

        // Final verification
        var pow = System.Math.Pow(1 + rate, nper);
        var finalCheck = System.Math.Abs(rate) < 1e-10
            ? pv + pmt * nper + fv
            : pv * pow + pmt * (1 + rate * type) * (pow - 1) / rate + fv;

        if (System.Math.Abs(finalCheck) > 0.01)
        {
            // Solution didn't converge well enough
            return CellValue.Error("#NUM!");
        }

        return CellValue.FromNumber(rate);
    }
}
