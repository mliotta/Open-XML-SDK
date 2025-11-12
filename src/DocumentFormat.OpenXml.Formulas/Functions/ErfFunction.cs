// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ERF function.
/// ERF(lower_limit, [upper_limit]) - returns the error function integrated between lower_limit and upper_limit.
/// </summary>
public sealed class ErfFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ErfFunction Instance = new();

    private ErfFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ERF";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length < 1 || args.Length > 2)
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

        var lowerLimit = args[0].NumericValue;

        if (args.Length == 1)
        {
            // ERF(x) = erf(x) - erf(0) = erf(x)
            return CellValue.FromNumber(ErrorFunction(lowerLimit));
        }

        if (args[1].IsError)
        {
            return args[1];
        }

        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }

        var upperLimit = args[1].NumericValue;

        // ERF(x, y) = erf(y) - erf(x)
        return CellValue.FromNumber(ErrorFunction(upperLimit) - ErrorFunction(lowerLimit));
    }

    /// <summary>
    /// Computes the error function using an approximation.
    /// Based on Abramowitz and Stegun formula 7.1.26.
    /// </summary>
    private static double ErrorFunction(double x)
    {
        // Constants for the approximation
        const double a1 = 0.254829592;
        const double a2 = -0.284496736;
        const double a3 = 1.421413741;
        const double a4 = -1.453152027;
        const double a5 = 1.061405429;
        const double p = 0.3275911;

        // Save the sign of x
        int sign = x < 0 ? -1 : 1;
        x = System.Math.Abs(x);

        // A&S formula 7.1.26
        double t = 1.0 / (1.0 + p * x);
        double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * System.Math.Exp(-x * x);

        return sign * y;
    }
}
