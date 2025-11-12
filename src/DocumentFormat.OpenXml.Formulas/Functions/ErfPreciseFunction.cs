// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the ERF.PRECISE function.
/// ERF.PRECISE(x) - returns the error function integrated between 0 and x.
/// </summary>
public sealed class ErfPreciseFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ErfPreciseFunction Instance = new();

    private ErfPreciseFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "ERF.PRECISE";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 1)
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

        var x = args[0].NumericValue;

        return CellValue.FromNumber(ErrorFunction(x));
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
