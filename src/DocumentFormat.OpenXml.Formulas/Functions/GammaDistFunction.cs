// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAMMA.DIST function.
/// GAMMA.DIST(x, alpha, beta, cumulative) - returns the gamma distribution.
/// </summary>
public sealed class GammaDistFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GammaDistFunction Instance = new();

    private GammaDistFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAMMA.DIST";

    /// <inheritdoc/>
    public FormulaResult Execute(CellContext context, FormulaResult[] args)
    {
        if (args.Length != 4)
        {
            return FormulaResult.Error("#VALUE!");
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
        if (args[0].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double x = args[0].NumericValue;

        if (x < 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Get alpha (shape parameter)
        if (args[1].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double alpha = args[1].NumericValue;

        if (alpha <= 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Get beta (scale parameter)
        if (args[2].Type != FormulaResultType.Number)
        {
            return FormulaResult.Error("#VALUE!");
        }
        double beta = args[2].NumericValue;

        if (beta <= 0)
        {
            return FormulaResult.Error("#NUM!");
        }

        // Get cumulative flag
        bool cumulative;
        if (args[3].Type == FormulaResultType.Boolean)
        {
            cumulative = args[3].BoolValue;
        }
        else if (args[3].Type == FormulaResultType.Number)
        {
            cumulative = args[3].NumericValue != 0;
        }
        else
        {
            return FormulaResult.Error("#VALUE!");
        }

        try
        {
            double result;
            if (cumulative)
            {
                // CDF: Regularized incomplete gamma function
                result = StatisticalHelper.GammaCDF(x / beta, alpha);
            }
            else
            {
                // PDF: (x^(alpha-1) * exp(-x/beta)) / (beta^alpha * Gamma(alpha))
                if (x == 0.0)
                {
                    if (alpha < 1.0)
                        return FormulaResult.FromNumber(double.PositiveInfinity);
                    else if (alpha == 1.0)
                        return FormulaResult.FromNumber(1.0 / beta);
                    else
                        return FormulaResult.FromNumber(0.0);
                }

                double logPdf = (alpha - 1.0) * System.Math.Log(x) - x / beta -
                               alpha * System.Math.Log(beta) - StatisticalHelper.LogGamma(alpha);
                result = System.Math.Exp(logPdf);
            }

            return FormulaResult.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return FormulaResult.Error("#NUM!");
        }
    }
}
