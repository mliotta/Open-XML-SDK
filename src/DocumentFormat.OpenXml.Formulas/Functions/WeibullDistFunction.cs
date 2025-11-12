// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the WEIBULL.DIST function.
/// WEIBULL.DIST(x, alpha, beta, cumulative) - returns the Weibull distribution.
/// </summary>
public sealed class WeibullDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly WeibullDistFunction Instance = new();

    private WeibullDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "WEIBULL.DIST";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 4)
        {
            return CellValue.Error("#VALUE!");
        }

        // Check for errors in arguments
        foreach (var arg in args)
        {
            if (arg.IsError)
            {
                return arg;
            }
        }

        // Get x value
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double x = args[0].NumericValue;

        if (x < 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get alpha (shape parameter)
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double alpha = args[1].NumericValue;

        if (alpha <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get beta (scale parameter)
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double beta = args[2].NumericValue;

        if (beta <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get cumulative flag
        bool cumulative;
        if (args[3].Type == CellValueType.Boolean)
        {
            cumulative = args[3].BoolValue;
        }
        else if (args[3].Type == CellValueType.Number)
        {
            cumulative = args[3].NumericValue != 0;
        }
        else
        {
            return CellValue.Error("#VALUE!");
        }

        double result;
        if (cumulative)
        {
            // CDF: 1 - exp(-(x/beta)^alpha)
            result = 1.0 - System.Math.Exp(-System.Math.Pow(x / beta, alpha));
        }
        else
        {
            // PDF: (alpha/beta) * (x/beta)^(alpha-1) * exp(-(x/beta)^alpha)
            if (x == 0.0)
            {
                if (alpha < 1.0)
                    return CellValue.FromNumber(double.PositiveInfinity);
                else if (alpha == 1.0)
                    return CellValue.FromNumber(alpha / beta);
                else
                    return CellValue.FromNumber(0.0);
            }

            double ratio = x / beta;
            result = (alpha / beta) * System.Math.Pow(ratio, alpha - 1.0) * System.Math.Exp(-System.Math.Pow(ratio, alpha));
        }

        return CellValue.FromNumber(result);
    }
}
