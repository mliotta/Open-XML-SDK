// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the GAMMA.INV function.
/// GAMMA.INV(probability, alpha, beta) - returns the inverse of the gamma cumulative distribution function.
/// </summary>
public sealed class GammaInvFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly GammaInvFunction Instance = new();

    private GammaInvFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "GAMMA.INV";

    /// <inheritdoc/>
    public CellValue Execute(CellContext context, CellValue[] args)
    {
        if (args.Length != 3)
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

        // Get probability
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double probability = args[0].NumericValue;

        if (probability <= 0.0 || probability >= 1.0)
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

        try
        {
            // Use Newton-Raphson to find x such that GammaCDF(x/beta, alpha) = probability
            double x = alpha * beta; // Initial guess (mean of gamma distribution)

            for (int i = 0; i < 20; i++)
            {
                double cdf = StatisticalHelper.GammaCDF(x / beta, alpha);

                // PDF: (x^(alpha-1) * exp(-x/beta)) / (beta^alpha * Gamma(alpha))
                double logPdf = (alpha - 1.0) * System.Math.Log(x) - x / beta -
                               alpha * System.Math.Log(beta) - StatisticalHelper.LogGamma(alpha);
                double pdf = System.Math.Exp(logPdf);

                if (System.Math.Abs(pdf) < 1e-20)
                    break;

                double delta = (cdf - probability) / pdf;
                x -= delta;

                if (x < 0.0001) x = 0.0001;

                if (System.Math.Abs(delta) < 1e-8)
                    break;
            }

            return CellValue.FromNumber(x);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
