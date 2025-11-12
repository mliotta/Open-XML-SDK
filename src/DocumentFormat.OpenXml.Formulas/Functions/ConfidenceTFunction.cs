// Copyright (c) Matt Liotta
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features.FormulaEvaluation.Compilation;

namespace DocumentFormat.OpenXml.Features.FormulaEvaluation.Functions;

/// <summary>
/// Implements the CONFIDENCE.T function.
/// CONFIDENCE.T(alpha, standard_dev, size) - returns the confidence interval for a population mean using Student's t-distribution.
/// </summary>
public sealed class ConfidenceTFunction : IFunctionImplementation
{
    /// <summary>
    /// Gets the singleton instance.
    /// </summary>
    public static readonly ConfidenceTFunction Instance = new();

    private ConfidenceTFunction()
    {
    }

    /// <inheritdoc/>
    public string Name => "CONFIDENCE.T";

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

        // Get alpha (significance level)
        if (args[0].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double alpha = args[0].NumericValue;

        if (alpha <= 0 || alpha >= 1)
        {
            return CellValue.Error("#NUM!");
        }

        // Get standard deviation
        if (args[1].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double standardDev = args[1].NumericValue;

        if (standardDev <= 0)
        {
            return CellValue.Error("#NUM!");
        }

        // Get sample size
        if (args[2].Type != CellValueType.Number)
        {
            return CellValue.Error("#VALUE!");
        }
        double size = args[2].NumericValue;

        if (size < 1)
        {
            return CellValue.Error("#NUM!");
        }

        int df = (int)size - 1; // degrees of freedom

        try
        {
            // Confidence interval = t * (stdev / sqrt(n))
            // where t is the critical value from Student's t-distribution
            // Using approximation for t-inverse (simplified)
            // For large n, t-distribution approaches normal distribution
            double t;
            if (df >= 30)
            {
                // Use normal approximation for large sample sizes
                t = StatisticalHelper.NormSInv(1 - alpha / 2);
            }
            else
            {
                // Simplified t-value approximation
                // This is a basic approximation; a full implementation would use proper t-distribution inverse
                double z = StatisticalHelper.NormSInv(1 - alpha / 2);

                // Adjust z-score to approximate t-value using a simple correction
                // t ≈ z * sqrt((df + 1) / df) for small samples
                t = z * System.Math.Sqrt((df + 1.0) / df);
            }

            double result = t * (standardDev / System.Math.Sqrt(size));
            return CellValue.FromNumber(result);
        }
        catch (System.ArgumentException)
        {
            return CellValue.Error("#NUM!");
        }
    }
}
